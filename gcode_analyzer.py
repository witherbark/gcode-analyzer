#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gcode_analyzer.py —— 3D 打印 G 代码分析工具
================================================================
用途
    读取切片软件生成的 .gcode 文件，输出一组可用于工艺分析的量化指标：
    路径长度、耗材消耗、成本、逐层统计、温度曲线、风扇策略等。

设计动机
    切片软件只给出结果，不解释数字来源，也无法批量对比不同参数方案的代价。
    本工具把 G 代码当作结构化数据解析，让耗材、时间、成本三项指标从
    经验判断转为可计算、可复现的量化依据，服务于报价、参数调优与工艺归档。

代码结构（从下往上读也行，从 main() 开始读最好）
    find_files()      找到要分析的 gcode 文件
    analyze_file()    核心：逐行解析一个 gcode 文件
    print_report()    把结果打印成人能看的报告
    build_figure()    画图（需要 matplotlib，没装就自动跳过）
    main()            命令行入口

作者备注：本脚本所有度量单位为 mm / g / 分钟，成本单位为元。
          默认耗材按 PLA 1.75mm 计算，可用命令行参数覆盖。
================================================================
"""

# ---- 标准库导入 ----------------------------------------------------------
# argparse: 解析命令行参数，支持 --density 1.24 这类覆盖项
# json:     把结果导出成 JSON，方便别的程序或 Excel 读取
# math:     算圆面积要用到 pi
# re:       正则表达式，用来从注释里抠出"预计打印时间"
# sys:      没有输入文件时打印提示并退出
# csv:      导出逐层数据，方便丢进 Excel 做图表
import argparse
import csv
import json
import math
import re
import sys
from pathlib import Path

# ---- 常量：常用耗材密度表 (g/cm³) ---------------------------------------
# 工艺计算常用数据，用于把"挤出长度"换算为"克数"
MATERIAL_DENSITY = {
    "PLA": 1.24,
    "PETG": 1.27,
    "ABS": 1.04,
    "ASA": 1.07,
    "TPU": 1.21,
    "NYLON": 1.14,
    "PC": 1.20,
}

# ---- 正则：从 gcode 里提取各种信息 ---------------------------------------
# G 代码里一行的格式通常是：  命令 参数   ; 注释
# 比如:  G1 X10.5 Y20.0 E0.85 F1800 ; 打印外壁
RE_LINE = re.compile(r"^\s*([GMgm]\d+)\s*(.*)$")          # 抓命令本身，如 G1 / M104
RE_PARAM = re.compile(r"([A-Za-z])\s*(-?\d*\.?\d+)")       # 抓参数，如 X10.5 E0.85

# 时间信息：不同切片软件写的格式不一样，所以准备好几种表达式
RE_TIME_SEC = re.compile(r";\s*TIME\s*:\s*(\d+)", re.I)                    # ;TIME:3720
RE_TIME_HMS = re.compile(r"(\d+)\s*h\s*(\d+)\s*m(?:\s*(\d+)\s*s)?", re.I)   # 1h 2m 3s
RE_TIME_MS = re.compile(r";\s*estimated printing time[^=]*=\s*(.+)$", re.I)  # ; estimated printing time = 1h 2m

# 层变化标记：PrusaSlicer/OrcaSlicer 用 ;LAYER_CHANGE，Cura 用 ;LAYER:3
RE_LAYER_CHANGE = re.compile(r";\s*LAYER_CHANGE", re.I)
RE_LAYER_CURA = re.compile(r";\s*LAYER\s*:\s*(\d+)", re.I)

# 机型信息
RE_PRINTER = re.compile(r";\s*(?:printer_model|model)\s*=\s*(.+)$", re.I)
RE_FILAMENT = re.compile(r";\s*filament_type\s*=\s*(.+)$", re.I)
RE_LAYER_HEIGHT = re.compile(r";\s*layer_height\s*=\s*([\d.]+)", re.I)
RE_NOZZLE = re.compile(r";\s*nozzle_diameter\s*=\s*([\d.]+)", re.I)


# =========================================================================
# 第 1 部分：辅助函数
# =========================================================================

def make_state():
    """
    创建一个"状态字典"，用来在逐行解析时记录当前位置和累计量。
    用一个 dict 而不是几十个变量，是为了让 analyze_file() 读起来清爽。
    """
    return {
        # --- 当前坐标 ---
        "x": 0.0, "y": 0.0, "z": 0.0, "e": 0.0,
        "f": 0.0,                 # 当前进给速度 mm/min
        # --- 挤出模式 ---
        "relative_e": False,      # M83 之后为 True（相对挤出），M82 之后为 False
        # --- 累计量 ---
        "total_move": 0.0,        # 空中移动总长 (mm)
        "total_extrude": 0.0,     # 挤出移动总长 (mm)
        "total_e": 0.0,           # 挤出耗材总长 (mm)  ← 算重量用这个
        "total_time_s": 0.0,      # 按 F 值累计的估算时间 (秒)
        "segment_count": 0,       # 移动指令条数
        # --- 温度 / 风扇 ---
        "nozzle_temps": [],       # [(行号, 温度)] 记录 M104/M109
        "bed_temps": [],          # [(行号, 温度)] 记录 M140/M190
        "fan_values": [],         # [(行号, 风扇PWM 0-255)] 记录 M106
        # --- 逐层统计 ---
        "layers": [],             # 每层一个 dict
        "current_layer": None,    # 正在累计的层
        "layer_index": -1,
        # --- 元数据 ---
        "printer": None,
        "filament_type": None,
        "layer_height": None,
        "nozzle_diameter": None,
        "estimated_time_from_comment": None,   # 切片软件自己写的预计时间(分钟)
    }


def parse_line(line):
    """
    把一行 gcode 拆成 (命令, {参数: 值})。
    识别不出来就返回 (None, {})。

    例子:
        'G1 X10.5 E0.85 F1800'  ->  ('G1', {'X': 10.5, 'E': 0.85, 'F': 1800.0})
    """
    # 先去掉注释部分（分号后面的内容），但注释在别处单独处理
    code_part = line.split(";", 1)[0].strip()
    if not code_part:
        return None, {}

    m = RE_LINE.match(code_part)
    if not m:
        return None, {}

    command = m.group(1).upper()
    params = {}
    for letter, value in RE_PARAM.findall(m.group(2)):
        params[letter.upper()] = float(value)
    return command, params


def parse_time_comment(text):
    """
    从一行注释里解析出"预计打印时间"，统一转成分钟(float)。
    解析不出来返回 None。
    """
    # 格式一: ;TIME:3720   （单位是秒）
    m = RE_TIME_SEC.search(text)
    if m:
        return int(m.group(1)) / 60.0

    # 格式二/三: 1h 2m 3s  或  1h 2m
    m = RE_TIME_HMS.search(text)
    if m:
        hours = int(m.group(1))
        minutes = int(m.group(2))
        seconds = int(m.group(3)) if m.group(3) else 0
        return hours * 60 + minutes + seconds / 60.0

    return None


# =========================================================================
# 第 2 部分：核心 —— 逐行解析
# =========================================================================

def analyze_file(path, density=1.24, filament_diameter=1.75,
                 price_per_kg=60.0, machine_rate_per_hour=2.0, labor_cost=0.0):
    """
    解析一个 gcode 文件，返回一个结果字典。

    参数说明
        density              耗材密度 g/cm³（PLA 默认 1.24）
        filament_diameter    耗材直径 mm（绝大多数桌面机是 1.75）
        price_per_kg         耗材单价 元/kg
        machine_rate_per_hour 机器工时费 元/小时（电费+设备折旧，自己估一个）
        labor_cost           固定的人工/后处理成本 元（比如拆支撑、打磨）
    """
    state = make_state()
    total_lines = 0

    # 用 utf-8 读，遇到非法字节就忽略（不同切片软件编码不一致，容错更稳）
    with open(path, "r", encoding="utf-8", errors="ignore") as fh:
        for lineno, raw_line in enumerate(fh, start=1):
            total_lines = lineno
            line = raw_line.rstrip("\n")

            # ---------- 第一步：处理注释里的元数据 ----------
            if ";" in line:
                comment = line[line.index(";"):]

                # 机型 / 材料 / 层高 / 喷嘴（只取第一次出现）
                if state["printer"] is None:
                    m = RE_PRINTER.search(comment)
                    if m:
                        state["printer"] = m.group(1).strip()
                if state["filament_type"] is None:
                    m = RE_FILAMENT.search(comment)
                    if m:
                        state["filament_type"] = m.group(1).strip().upper()
                if state["layer_height"] is None:
                    m = RE_LAYER_HEIGHT.search(comment)
                    if m:
                        state["layer_height"] = float(m.group(1))
                if state["nozzle_diameter"] is None:
                    m = RE_NOZZLE.search(comment)
                    if m:
                        state["nozzle_diameter"] = float(m.group(1))

                # 切片软件自报的预计时间
                if state["estimated_time_from_comment"] is None:
                    t = parse_time_comment(comment)
                    if t is not None:
                        state["estimated_time_from_comment"] = t

                # 层变化（PrusaSlicer / OrcaSlicer 风格）
                if RE_LAYER_CHANGE.search(comment):
                    _start_new_layer(state)

                # 层变化（Cura 风格 ;LAYER:3）
                m = RE_LAYER_CURA.search(comment)
                if m:
                    _start_new_layer(state, index=int(m.group(1)))

            # ---------- 第二步：解析命令 ----------
            command, params = parse_line(line)
            if command is None:
                continue

            # --- 挤出模式切换 ---
            if command == "M82":
                state["relative_e"] = False
                continue
            if command == "M83":
                state["relative_e"] = True
                continue

            # --- 温度 ---
            if command in ("M104", "M109") and "S" in params:
                state["nozzle_temps"].append((lineno, params["S"]))
                continue
            if command in ("M140", "M190") and "S" in params:
                state["bed_temps"].append((lineno, params["S"]))
                continue

            # --- 风扇 ---
            if command == "M106":
                state["fan_values"].append((lineno, params.get("S", 255)))
                continue

            # --- 只关心直线移动 G0 / G1 ---
            if command not in ("G0", "G1"):
                continue

            # 1) 更新进给速度（F 值单位是 mm/min）
            if "F" in params:
                state["f"] = params["F"]

            # 2) 取出这一条指令的目标坐标（没写就沿用当前值）
            nx = params.get("X", state["x"])
            ny = params.get("Y", state["y"])
            nz = params.get("Z", state["z"])

            # 3) 算这一段走了多远（三维距离）
            dist = math.sqrt((nx - state["x"]) ** 2 +
                             (ny - state["y"]) ** 2 +
                             (nz - state["z"]) ** 2)

            # 4) 算这一段挤出了多少耗材
            #    绝对模式(M82)：E 是"累计挤出总量"，所以要用差值
            #    相对模式(M83)：E 就是"本段挤出量"，直接用
            if "E" in params:
                if state["relative_e"]:
                    e_delta = params["E"]
                else:
                    e_delta = params["E"] - state["e"]
                    state["e"] = params["E"]
            else:
                e_delta = 0.0

            # 5) 累计
            state["segment_count"] += 1
            if e_delta > 0:
                state["total_extrude"] += dist
                state["total_e"] += e_delta
            else:
                state["total_move"] += dist

            # 6) 用 F 值估算这一段耗时：时间(秒) = 距离 / 速度(mm/min) * 60
            if state["f"] > 0 and dist > 0:
                seg_time = dist / state["f"] * 60.0
                state["total_time_s"] += seg_time
                if state["current_layer"] is not None:
                    state["current_layer"]["time_s"] += seg_time
                    state["current_layer"]["extrude_mm"] += max(e_delta, 0.0)

            # 7) 更新当前位置
            state["x"], state["y"], state["z"] = nx, ny, nz

    # ---------- 第三步：结算 ----------
    result = _summarize(
        path=path,
        state=state,
        total_lines=total_lines,
        density=density,
        filament_diameter=filament_diameter,
        price_per_kg=price_per_kg,
        machine_rate_per_hour=machine_rate_per_hour,
        labor_cost=labor_cost,
    )
    return result


def _start_new_layer(state, index=None):
    """
    开始记录新的一层。
    如果上一层的记录还开着，就先收尾（补上层号）。
    """
    if state["current_layer"] is not None:
        state["layers"].append(state["current_layer"])

    state["layer_index"] = index if index is not None else state["layer_index"] + 1
    state["current_layer"] = {
        "layer": state["layer_index"],
        "z": state["z"],
        "time_s": 0.0,
        "extrude_mm": 0.0,
    }


def _summarize(path, state, total_lines, density, filament_diameter,
               price_per_kg, machine_rate_per_hour, labor_cost):
    """把解析过程中的原始累计量换算成人类可读的指标，并计算成本。"""
    # 把最后一层收尾
    if state["current_layer"] is not None:
        state["layers"].append(state["current_layer"])

    # ---------- 耗材重量 ----------
    # 挤出长度(mm) -> 体积(mm³) -> 质量(g)
    #   截面积 = π * (d/2)²        (mm²)
    #   体积   = 长度 * 截面积     (mm³)
    #   质量   = 体积 / 1000 * 密度 (g)   ← 因为 1 cm³ = 1000 mm³
    radius = filament_diameter / 2.0
    area = math.pi * radius * radius
    volume_mm3 = state["total_e"] * area
    weight_g = volume_mm3 / 1000.0 * density

    # ---------- 时间 ----------
    # 优先用切片软件自己算的时间（它考虑了加速度、Jerk，更准）
    # 没写就用我们自己按 F 值累加的估算值
    est_from_comment = state["estimated_time_from_comment"]
    if est_from_comment is not None:
        print_time_min = est_from_comment
        time_source = "切片软件注释（更准）"
    else:
        print_time_min = state["total_time_s"] / 60.0
        time_source = "按 F 值自算估算（可能偏快，未考虑加速度/Jerk）"

    # ---------- 成本 ----------
    material_cost = weight_g / 1000.0 * price_per_kg
    machine_cost = print_time_min / 60.0 * machine_rate_per_hour
    total_cost = material_cost + machine_cost + labor_cost

    # ---------- 逐层整理 ----------
    layers = []
    for item in state["layers"]:
        layers.append({
            "layer": item["layer"],
            "z_mm": round(item["z"], 3),
            "time_min": round(item["time_s"] / 60.0, 2),
            "extrude_mm": round(item["extrude_mm"], 1),
        })

    return {
        "file": str(path),
        "file_size_kb": round(Path(path).stat().st_size / 1024.0, 1),
        "total_lines": total_lines,
        "printer": state["printer"],
        "filament_type": state["filament_type"],
        "layer_height_setting": state["layer_height"],
        "nozzle_diameter": state["nozzle_diameter"],
        "layers_detected": len(layers),
        "segment_count": state["segment_count"],
        "extrude_path_mm": round(state["total_extrude"], 1),
        "travel_path_mm": round(state["total_move"], 1),
        "filament_used_mm": round(state["total_e"], 1),
        "filament_used_m": round(state["total_e"] / 1000.0, 3),
        "weight_g": round(weight_g, 2),
        "print_time_min": round(print_time_min, 2),
        "print_time_human": format_duration(print_time_min),
        "time_source": time_source,
        "nozzle_temp_range": _temp_range(state["nozzle_temps"]),
        "bed_temp_range": _temp_range(state["bed_temps"]),
        "fan_max_pwm": max([v for _, v in state["fan_values"]], default=None),
        "fan_change_count": len(state["fan_values"]),
        "cost": {
            "material_yuan": round(material_cost, 2),
            "machine_yuan": round(machine_cost, 2),
            "labor_yuan": round(labor_cost, 2),
            "total_yuan": round(total_cost, 2),
            "assumptions": {
                "density_g_cm3": density,
                "filament_diameter_mm": filament_diameter,
                "price_per_kg_yuan": price_per_kg,
                "machine_rate_per_hour_yuan": machine_rate_per_hour,
            },
        },
        "layers": layers,
    }


def _temp_range(entries):
    """从 [(行号, 温度)] 列表里取出温度范围，忽略 0（关机指令）。"""
    values = [t for _, t in entries if t > 0]
    if not values:
        return None
    return {"min": min(values), "max": max(values), "changes": len(values)}


def format_duration(minutes):
    """把分钟数格式化成 '2h 15m' 这种好读的形式。"""
    total_seconds = int(round(minutes * 60))
    hours, remainder = divmod(total_seconds, 3600)
    mins, secs = divmod(remainder, 60)
    if hours:
        return f"{hours}h {mins}m"
    if mins:
        return f"{mins}m {secs}s"
    return f"{secs}s"


# =========================================================================
# 第 3 部分：输出
# =========================================================================

def print_report(result):
    """将分析结果输出为格式化的可读报告（终端直接阅读）。"""
    c = result["cost"]
    line = "=" * 64
    print(line)
    print(f"G代码分析报告  |  {Path(result['file']).name}")
    print(line)

    print("\n【文件信息】")
    print(f"  文件大小      : {result['file_size_kb']} KB")
    print(f"  总行数        : {result['total_lines']:,}")
    print(f"  机型          : {result['printer'] or '未识别'}")
    print(f"  材料类型      : {result['filament_type'] or '未识别'}")
    print(f"  层高设定      : {result['layer_height_setting'] or '未识别'} mm")
    print(f"  喷嘴直径      : {result['nozzle_diameter'] or '未识别'} mm")

    print("\n【结构】")
    print(f"  识别层数      : {result['layers_detected']}")
    print(f"  移动指令条数  : {result['segment_count']:,}")

    print("\n【路径】")
    print(f"  挤出路径长度  : {result['extrude_path_mm']:,.1f} mm")
    print(f"  空驶路径长度  : {result['travel_path_mm']:,.1f} mm")
    total_path = result["extrude_path_mm"] + result["travel_path_mm"]
    if total_path > 0:
        ratio = result["extrude_path_mm"] / total_path * 100
        print(f"  路径有效率    : {ratio:.1f}%  （挤出路径占比，越高说明空驶越少）")

    print("\n【耗材与时间】")
    print(f"  耗材长度      : {result['filament_used_m']} m")
    print(f"  耗材重量      : {result['weight_g']} g")
    print(f"  打印时间      : {result['print_time_human']}  ({result['print_time_min']} 分钟)")
    print(f"  时间来源      : {result['time_source']}")

    print("\n【温度与风扇】")
    nt = result["nozzle_temp_range"]
    bt = result["bed_temp_range"]
    if nt:
        print(f"  喷嘴温度      : {nt['min']}-{nt['max']} C （{nt['changes']} 次设定）")
    else:
        print("  喷嘴温度      : 未识别")

    if bt:
        print(f"  热床温度      : {bt['min']}-{bt['max']} C （{bt['changes']} 次设定）")
    else:
        print("  热床温度      : 未识别")
    print(f"  风扇最大PWM   : {result['fan_max_pwm'] if result['fan_max_pwm'] is not None else '未识别'}")
    print(f"  风扇调节次数  : {result['fan_change_count']}")

    print("\n【成本核算】")
    print(f"  耗材成本      : {c['material_yuan']} CNY")
    print(f"  机器工时      : {c['machine_yuan']} CNY")
    print(f"  人工/后处理   : {c['labor_yuan']} CNY")
    print("  " + "-" * 20)
    print(f"  单件总成本    : {c['total_yuan']} CNY")
    a = c["assumptions"]
    print(f"  (按 {a['filament_diameter_mm']}mm 耗材、密度 {a['density_g_cm3']} g/cm3、"
          f"{a['price_per_kg_yuan']} CNY/kg、工时 {a['machine_rate_per_hour_yuan']} CNY/h 计算)")

    # 逐层摘要：只打印前 5 层和最长的 3 层，避免刷屏
    if result["layers"]:
        print("\n【逐层统计】")
        print(f"  {'层号':>6} {'Z高度(mm)':>10} {'耗时(min)':>10} {'挤出(mm)':>12}")
        for item in result["layers"][:5]:
            print(f"  {item['layer']:>6} {item['z_mm']:>10.3f} "
                  f"{item['time_min']:>10.2f} {item['extrude_mm']:>12.1f}")
        if len(result["layers"]) > 8:
            print("  ...")
            slowest = sorted(result["layers"], key=lambda x: -x["time_min"])[:3]
            print("  最耗时的 3 层:")
            for item in slowest:
                print(f"  {item['layer']:>6} {item['z_mm']:>10.3f} "
                      f"{item['time_min']:>10.2f} {item['extrude_mm']:>12.1f}")

    print("\n" + line)


def export_layers_csv(result, out_path):
    """把逐层数据导出成 CSV，可以直接丢进 Excel 画图。"""
    with open(out_path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=["layer", "z_mm", "time_min", "extrude_mm"])
        writer.writeheader()
        writer.writerows(result["layers"])
    print(f"[已导出] 逐层数据 -> {out_path}")


def build_figure(result, out_path):
    """
    画两张图：耗材/时间沿 Z 高度的分布，以及温度设定变化。
    matplotlib 没安装就静默跳过 —— 不影响主功能。
    """
    try:
        import matplotlib
        matplotlib.use("Agg")           # 不弹窗，直接存文件
        import matplotlib.pyplot as plt
    except ImportError:
        print("[提示] 未安装 matplotlib，跳过绘图（不影响其他功能）")
        return False

    layers = result["layers"]
    if not layers:
        print("[提示] 没有逐层数据，跳过绘图")
        return False

    fig, axes = plt.subplots(1, 2, figsize=(12, 4.5))

    zs = [x["z_mm"] for x in layers]
    times = [x["time_min"] for x in layers]
    exts = [x["extrude_mm"] for x in layers]

    # 左图：逐层耗时
    axes[0].plot(times, zs, linewidth=1.6, color="#2563eb")
    axes[0].set_xlabel("Layer time (min)")
    axes[0].set_ylabel("Z height (mm)")
    axes[0].set_title("Time per layer")
    axes[0].grid(alpha=0.3)

    # 右图：逐层挤出量
    axes[1].plot(exts, zs, linewidth=1.6, color="#dc2626")
    axes[1].set_xlabel("Extrusion (mm)")
    axes[1].set_ylabel("Z height (mm)")
    axes[1].set_title("Extrusion per layer")
    axes[1].grid(alpha=0.3)

    fig.suptitle(f"G-code analysis: {Path(result['file']).name}", fontsize=11)
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    print(f"[已导出] 图表 -> {out_path}")
    return True


# =========================================================================
# 第 4 部分：命令行入口
# =========================================================================

def find_files(target):
    """target 可以是单个文件，也可以是文件夹（自动找里面所有 .gcode）。"""
    p = Path(target)
    if p.is_file():
        return [p]
    if p.is_dir():
        found = sorted(p.glob("*.gcode")) + sorted(p.glob("*.gco")) + sorted(p.glob("*.g"))
        return found
    return []


def build_parser():
    parser = argparse.ArgumentParser(
        description="3D 打印 G 代码分析工具：统计路径、耗材、时间与成本",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例
  # 分析单个文件
  python gcode_analyzer.py part.gcode

  # 分析整个文件夹，并导出 JSON + 逐层 CSV + 图表
  python gcode_analyzer.py ./gcode/ --json report.json --csv layers.csv --figure chart.png

  # 换成 PETG 并调整成本假设
  python gcode_analyzer.py part.gcode --material PETG --price 80 --machine-rate 3 --labor 2
        """,
    )
    parser.add_argument("input", help="要分析的 .gcode 文件或包含 gcode 的文件夹")
    parser.add_argument("--json", dest="json_out", help="把完整结果导出为 JSON")
    parser.add_argument("--csv", dest="csv_out", help="把逐层数据导出为 CSV")
    parser.add_argument("--figure", dest="figure_out", help="生成分析图表 PNG")
    parser.add_argument("--material", default="PLA", choices=sorted(MATERIAL_DENSITY.keys()),
                        help="耗材类型，用于查密度表（默认 PLA）")
    parser.add_argument("--density", type=float, default=None,
                        help="直接指定密度 g/cm³，会覆盖 --material")
    parser.add_argument("--diameter", type=float, default=1.75, help="耗材直径 mm（默认 1.75）")
    parser.add_argument("--price", type=float, default=60.0, help="耗材单价 元/kg（默认 60）")
    parser.add_argument("--machine-rate", type=float, default=2.0,
                        help="机器工时费 元/小时（默认 2.0）")
    parser.add_argument("--labor", type=float, default=0.0,
                        help="人工/后处理成本 元（默认 0）")
    parser.add_argument("--quiet", action="store_true", help="只输出 JSON，不打印报告")
    return parser


def main(argv=None):
    # Windows 控制台默认是 GBK 编码，直接 print 中文/特殊符号会报
    # UnicodeEncodeError。这里强制把标准输出切成 UTF-8，避免报告打不出来。
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    args = build_parser().parse_args(argv)

    files = find_files(args.input)
    if not files:
        print(f"[错误] 没找到任何 gcode 文件: {args.input}", file=sys.stderr)
        return 1

    density = args.density if args.density is not None else MATERIAL_DENSITY[args.material]

    all_results = []
    for path in files:
        result = analyze_file(
            path,
            density=density,
            filament_diameter=args.diameter,
            price_per_kg=args.price,
            machine_rate_per_hour=args.machine_rate,
            labor_cost=args.labor,
        )
        all_results.append(result)
        if not args.quiet:
            print_report(result)

    if args.json_out:
        payload = all_results[0] if len(all_results) == 1 else all_results
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
        print(f"[已导出] 完整结果 -> {args.json_out}")

    if args.csv_out and all_results:
        export_layers_csv(all_results[0], args.csv_out)

    if args.figure_out and all_results:
        build_figure(all_results[0], args.figure_out)

    return 0


if __name__ == "__main__":
    sys.exit(main())
