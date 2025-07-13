#!/usr/bin/env python3
"""
KML ➜ MSFS2020 .pln converter

用法:
    python kml2pln.py
    # 随后按提示输入 KML 文件路径
"""

import os
import re
import math
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom

NS = {"kml": "http://www.opengis.net/kml/2.2"}

def dec_to_dms(value: float, is_lat: bool) -> str:
    """十进制度 ➜ DMS 字符串（N/S, E/W）"""
    hem = ("N" if value >= 0 else "S") if is_lat else ("E" if value >= 0 else "W")
    value = abs(value)
    d = int(value)
    m_full = (value - d) * 60
    m = int(m_full)
    s = (m_full - m) * 60
    return f"{hem}{d}° {m}' {s:05.2f}\""

def alt_m_to_ft_str(alt_m: float) -> str:
    """米 ➜ 英尺（字符串格式 +000028.00）"""
    ft = int(round(alt_m * 3.28084))
    return f"+{ft:06d}.00"

def parse_kml(kml_path: str):
    """读取 KML，返回 (标题, 航点列表[ (name, lat, lon, alt_m) ])"""
    tree = ET.parse(kml_path)
    root = tree.getroot()

    title_el = root.find(".//kml:Document/kml:name", NS)
    title = title_el.text.strip() if title_el is not None else os.path.basename(kml_path)

    waypoints = []
    for pm in root.findall(".//kml:Placemark", NS):
        point = pm.find("kml:Point", NS)
        if point is None:
            continue                          # 跳过 LineString / 路径等
        coord_text = point.find("kml:coordinates", NS).text.strip()
        lon, lat, alt = map(float, coord_text.split(",")[:3])
        name = pm.find("kml:name", NS).text.strip()
        waypoints.append((name, lat, lon, alt))

    if not waypoints:
        raise ValueError("未在 KML 中找到任何航点 (Placemark with <Point>).")

    return title, waypoints

def build_pln_xml(title: str, wpts):
    """构造 SimBase.Document 根节点"""
    # 尝试从标题抓 ICAO
    m = re.match(r"(\w{4})\s+to\s+(\w{4})", title, re.I)
    dep_id, dest_id = (m.group(1).upper(), m.group(2).upper()) if m else (wpts[0][0], wpts[-1][0])

    root = ET.Element("SimBase.Document", Type="AceXML", version="1,0")
    ET.SubElement(root, "Descr").text = "AceXML Document"
    fp = ET.SubElement(root, "FlightPlan.FlightPlan")

    # 基本信息
    ET.SubElement(fp, "Title").text = title
    ET.SubElement(fp, "FPType").text = "IFR"
    ET.SubElement(fp, "RouteType").text = "Direct"
    ET.SubElement(fp, "CruisingAlt").text = "00000"

    # 出发 / 目的地
    lat0, lon0, alt0 = wpts[0][1:]
    latN, lonN, altN = wpts[-1][1:]
    ET.SubElement(fp, "DepartureID").text = dep_id
    ET.SubElement(fp, "DepartureLLA").text = f"{dec_to_dms(lat0, True)},{dec_to_dms(lon0, False)},{alt_m_to_ft_str(alt0)}"
    ET.SubElement(fp, "DestinationID").text = dest_id
    ET.SubElement(fp, "DestinationLLA").text = f"{dec_to_dms(latN, True)},{dec_to_dms(lonN, False)},{alt_m_to_ft_str(altN)}"

    ET.SubElement(fp, "Descr").text = f"{title} Created by KML2MSFS"
    ET.SubElement(fp, "DepartureName").text = dep_id
    ET.SubElement(fp, "DestinationName").text = dest_id

    app = ET.SubElement(fp, "AppVersion")
    ET.SubElement(app, "AppVersionMajor").text = "11"
    ET.SubElement(app, "AppVersionBuild").text = "282174"

    # 航路点
    for idx, (name, lat, lon, alt_m) in enumerate(wpts):
        wp = ET.SubElement(fp, "ATCWaypoint", id=name)
        if idx in (0, len(wpts) - 1):
            wp_type = "Airport"
        elif len(name) == 5 and name.isalpha():
            wp_type = "Intersection"
        else:
            wp_type = "User"

        ET.SubElement(wp, "ATCWaypointType").text = wp_type
        ET.SubElement(wp, "WorldPosition").text = (
            f"{dec_to_dms(lat, True)},{dec_to_dms(lon, False)},{alt_m_to_ft_str(alt_m)}"
        )

        # 写 ICAO 容器（对 Airport/Intersection 有用）
        if wp_type in ("Airport", "Intersection"):
            icao = ET.SubElement(wp, "ICAO")
            ET.SubElement(icao, "ICAOIdent").text = name[:4] if wp_type == "Airport" else name

    return root

def write_pln(root, out_path: str):
    """保存带缩进的 XML"""
    rough = ET.tostring(root, encoding="utf-8")
    pretty = minidom.parseString(rough).toprettyxml(indent="    ", encoding="utf-8")
    with open(out_path, "wb") as f:
        f.write(pretty)

def main():
    kml_path = input("KML 文件路径: ").strip().strip('"').strip("'")
    if not os.path.isfile(kml_path):
        print("找不到文件:", kml_path)
        return

    title, waypoints = parse_kml(kml_path)
    pln_root = build_pln_xml(title, waypoints)
    out_path = os.path.splitext(kml_path)[0] + "_MSFS2020.pln"
    write_pln(pln_root, out_path)
    print("✅ 已生成:", out_path)

if __name__ == "__main__":
    main()
