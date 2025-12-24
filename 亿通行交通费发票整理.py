import os
import re
from pathlib import Path
import fitz  # PyMuPDF
import datetime
import sys

def get_invoice_number_from_pdf(pdf_path):
    """从发票PDF中提取发票号码"""
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()

        # 查找发票号码的常见模式
        # 通常发票号码会标注为"发票号码"或"NO."等
        patterns = [
            r'发票号码[：:]\s*([A-Z0-9]{8,20})',  # 发票号码: 后面跟着字母数字
            r'发票号[：:]\s*([A-Z0-9]{8,20})',    # 发票号: 后面跟着字母数字
            r'NO\.?\s*([A-Z0-9]{8,20})',         # NO. 后面跟着字母数字
            r'(\d{18,20})',                      # 20位左右的数字串（可能是发票号码）
            r'(\d{8,12})'                        # 8-12位数字串（传统发票号码）
        ]

        for pattern in patterns:
            matches = re.findall(pattern, text)
            if matches:
                # 排除可能是统一社会信用代码的91开头的18位数字
                for match in matches:
                    if len(match) == 18 and match.startswith('91'):
                        continue  # 跳过统一社会信用代码
                    return match
                # 如果只有统一社会信用代码，则返回未知发票号码
                return "未知发票号码"

    except Exception as e:
        print(f"警告 (pymupdf): 提取发票号码时处理发票 {pdf_path} 出错: {e}")

    return "未知发票号码"


def get_total_from_invoice_definitive(pdf_path):
    """遍历PDF的文本块，找到包含汉字大写金额的块，然后从该块中提取数字金额"""
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            blocks = page.get_text("blocks")
            for block in blocks:
                block_text = block[4]
                if '圆' in block_text or '整' in block_text:
                    match = re.search(r'([\d,]+\.\d{2})', block_text)
                    if match:
                        doc.close()
                        return float(match.group(1).replace(',', ''))
        doc.close()
    except Exception as e:
        print(f"警告 (pymupdf): 处理发票 {pdf_path} 时出错: {e}")
    return 0.0

def get_trip_data_definitive(pdf_path):
    """
    终极版行程解析函数，可以处理多种PDF文本布局。
    使用更精确的文本行解析方法处理公交行程单。
    返回 (行程列表, 摘要总额)。
    """
    trips = []
    summary_total = 0.0

    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page in doc:
            full_text += page.get_text()

        # 首先，从全文中获取摘要总额
        summary_match = re.search(r'合计\s*([\d,]+\.\d{2})\s*元', full_text)
        if summary_match:
            summary_total = float(summary_match.group(1).replace(',', ''))

        # 检查是否包含公交行程单的标识
        has_bus_format = '行程站点' in full_text and '金额(元)' in full_text

        if has_bus_format:
            # 使用更精确的解析方法处理公交行程单
            # 获取所有文本行以进行更精确的解析
            full_text_lines = []
            for page in doc:
                blocks = page.get_text("dict")["blocks"]
                for block in blocks:
                    if "lines" in block:
                        for line in block["lines"]:
                            line_text = ""
                            for span in line["spans"]:
                                line_text += span["text"]
                            if line_text.strip():
                                full_text_lines.append(line_text.strip())

            # 解析公交行程单格式
            i = 0
            while i < len(full_text_lines):
                line = full_text_lines[i]

                # 检查是否是序号行（只包含数字）
                if re.match(r'^\d+\s*$', line):
                    # 检查后续行是否包含日期
                    if i + 1 < len(full_text_lines):
                        date_line = full_text_lines[i + 1]
                        date_match = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', date_line)

                        if date_match:
                            date_str = date_match.group(1)

                            # 检查后续行是否包含时间
                            if i + 2 < len(full_text_lines):
                                time_line = full_text_lines[i + 2]
                                time_match = re.search(r'\d{1,2}:\d{2}-\d{1,2}:\d{2}', time_line)

                                if time_match:
                                    # 检查后续行是否包含站点和金额
                                    # 通常站点信息可能跨越一行或多行
                                    if i + 3 < len(full_text_lines):
                                        station_line1 = full_text_lines[i + 3]

                                        if i + 4 < len(full_text_lines):
                                            potential_station2_or_amount = full_text_lines[i + 4]

                                            # 检查第5行是否是金额
                                            amount_match = re.search(r'(\d+\.\d{2})', potential_station2_or_amount)
                                            if amount_match:
                                                # 第4行是站点1，第5行是金额
                                                station_info = station_line1
                                                amount = float(amount_match.group(1))

                                                trips.append({
                                                    'date': date_str,
                                                    'departure': station_info,
                                                    'destination': '',
                                                    'amount': amount
                                                })
                                                i += 5  # 跳过这5行
                                                continue
                                            elif i + 5 < len(full_text_lines):
                                                # 检查第6行是否是金额（站点跨越两行的情况）
                                                potential_amount_line = full_text_lines[i + 5]
                                                amount_match = re.search(r'(\d+\.\d{2})', potential_amount_line)
                                                if amount_match:
                                                    # 第4行和第5行都是站点信息，第6行是金额
                                                    station_info = station_line1 + potential_station2_or_amount
                                                    amount = float(amount_match.group(1))

                                                    trips.append({
                                                        'date': date_str,
                                                        'departure': station_info,
                                                        'destination': '',
                                                        'amount': amount
                                                    })
                                                    i += 6  # 跳过这6行
                                                    continue

                                # 如果时间匹配失败，但站点行包含金额，处理这种情况
                                amount_match = re.search(r'(\d+\.\d{2})', station_line1)
                                if amount_match:
                                    amount = float(amount_match.group(1))
                                    # 提取站点信息（去除金额部分）
                                    station_info = re.sub(r'\s*\d+\.\d{2}\s*$', '', station_line1).strip()

                                    trips.append({
                                        'date': date_str,
                                        'departure': station_info,
                                        'destination': '',
                                        'amount': amount
                                    })
                                    i += 4  # 跳过这4行
                                    continue
                i += 1
        else:
            # 处理非公交行程单（地铁等）
            for page in doc:
                blocks = page.get_text("blocks")
                for block in blocks:
                    block_text = block[4].strip()

                    # 尝试将一个块视为一个潜在的行程记录
                    # 规则：如果一个块包含 "年" "月" "日" 和一个金额，就尝试解析
                    if '年' in block_text and '月' in block_text and '日' in block_text and re.search(r'\d+\.\d{2}', block_text):
                        lines = block_text.split('\n')
                        # 如果块内少于3行，不太可能是完整的行程记录
                        if len(lines) < 3:
                            continue

                        try:
                            # 这是一个单行记录，用正则表达式匹配
                            if len(lines) == 1:
                                 trip_line_pattern = re.compile(r'^\d+\s+(\d{4}年\d{1,2}月\d{1,2}日)\s+[\d:-]+\s+(.+?)\s+([\d,]+\.\d{2})$')
                                 match = trip_line_pattern.match(lines[0])
                                 if match:
                                     date_str, station_info, amount_str = match.groups()
                                     amount = float(amount_str.replace(',', ''))
                                     trips.append({'date': date_str, 'departure': station_info, 'destination': '', 'amount': amount})
                            # 这是一个多行记录 (像19/trip.pdf)
                            else:
                                # 假设金额是最后一行
                                amount_line = [line for line in lines if re.search(r'(\d+\.\d{2})', line)]
                                if amount_line:
                                    amount = float(re.search(r'(\d+\.\d{2})', amount_line[-1]).group(1))
                                else:
                                    continue

                                date_str = ""
                                # 找到包含年份的行作为日期
                                for line in lines:
                                    if '年' in line:
                                        date_str = line
                                        break
                                # 将日期和金额之外的行合并为站点信息
                                station_info = " ".join([line for line in lines if date_str not in line and not re.search(r'\d+\.\d{2}', line) and not line.isdigit() and not re.match(r'[\d:-]+$', line)])

                                trips.append({'date': date_str, 'departure': station_info, 'destination': '', 'amount': amount})

                        except (ValueError, IndexError):
                            # 如果解析失败，就跳过这个块
                            continue
        doc.close()
        # 如果以上策略失败，使用pypdf作为备用方案进行最后尝试
        if not trips:
            from pypdf import PdfReader
            reader = PdfReader(pdf_path)
            pypdf_text = ""
            for page in reader.pages:
                pypdf_text += page.extract_text() + "\n"

            raw_lines = pypdf_text.split('\n')
            merged_lines = []
            for line in raw_lines:
                line = line.strip()
                if not line: continue
                if re.match(r'^\d+\s+\d{4}年', line):
                    merged_lines.append(line)
                elif merged_lines:
                    merged_lines[-1] += " " + line

            trip_line_pattern = re.compile(r'^\d+\s+(\d{4}年\d{1,2}月\d{1,2}日)\s+[\d:-]+\s+(.+?)\s+([\d,]+\.\d{2})$')
            for line in merged_lines:
                match = trip_line_pattern.match(line.strip())
                if match:
                    date_str, station_info, amount_str = match.groups()
                    amount = float(amount_str.replace(',', ''))
                    trips.append({'date': date_str, 'departure': station_info, 'destination': '', 'amount': amount})

        return trips, summary_total
    except Exception as e:
        print(f"警告: 处理行程单 {pdf_path} 时发生严重错误: {e}")
        return [], 0.0


def main():
    # 检查命令行参数
    if len(sys.argv) > 1:
        directory_name = sys.argv[1]
    else:
        directory_name = "发票"  # 默认目录名
    
    print(f"使用目录名称: {directory_name}")
    
    base_path = Path(f'./{directory_name}')  # 使用用户输入的目录名
    if not base_path.exists():
        print(f"错误: 目录 {base_path} 不存在")
        return

    invoice_dirs = sorted([d for d in base_path.iterdir() if d.is_dir() and d.name.isdigit()], key=lambda x: int(x.name))

    comparison_results = []
    all_trips = []

    print(f"处理{directory_name}中的发票和行程单...")
    print("=" * 80)

    for directory in invoice_dirs:
        trip_path = directory / "trip.pdf"
        invoice_path = directory / "invoice.pdf"

        status = ""
        parsed_trips, trip_total_summary = get_trip_data_definitive(trip_path)

        if parsed_trips:
            all_trips.extend(parsed_trips)
        elif trip_total_summary > 0:
            # 如果解析失败，但摘要金额存在，则将摘要金额作为一条记录
            all_trips.append({
                'date': '见摘要',
                'departure': f"来自 {directory.name}/trip.pdf (解析失败)",
                'destination': "",
                'amount': trip_total_summary,
            })

        invoice_total = get_total_from_invoice_definitive(invoice_path) if invoice_path.exists() else 0.0

        # 提取发票号码（从发票PDF内容中）
        invoice_number = get_invoice_number_from_pdf(invoice_path) if invoice_path.exists() else "未找到发票"

        if abs(trip_total_summary - invoice_total) < 0.01:
            status = "匹配"
        else:
            status = f"不匹配 (差额: {trip_total_summary - invoice_total:.2f})"

        comparison_results.append({
            'dir': directory.name,
            'invoice_number': invoice_number,
            'trip_total': trip_total_summary,
            'invoice_total': invoice_total,
            'status': status
        })
        print(f"处理目录: {directory_name}/{directory.name} -> 行程单: {trip_total_summary:.2f}, 发票: {invoice_total:.2f}, 发票号码: {invoice_number}, 状态: {status}")

    # 根据目录名称生成输出文件名
    output_filename = f'{directory_name}汇总.md'
    with open(output_filename, 'w', encoding='utf-8-sig') as f:
        f.write(f"# 2025年{directory_name}车票报销明细\n\n")

        f.write("## 第一部分: 金额比对摘要\n\n")
        f.write("| 目录 | 发票号码 | 行程单总额 | 发票总额 | 状态 |\n")
        f.write("|------|--------|----------|--------|------|\n")
        for res in comparison_results:
            f.write(f"| {res['dir']} | {res['invoice_number']} | {res['trip_total']:.2f} | {res['invoice_total']:.2f} | {res['status']} |\n")

        grand_total_trip = sum(res['trip_total'] for res in comparison_results)
        grand_total_invoice = sum(res['invoice_total'] for res in comparison_results)
        f.write(f"| 总计 | | {grand_total_trip:.2f} | {grand_total_invoice:.2f} | |\n\n")

        f.write("## 第二部分: 详细行程清单\n\n")

        all_trips.sort(key=lambda x: x.get('date', ''))

        grand_total_detailed = sum(trip['amount'] for trip in all_trips)
        f.write(f"**总计行程: {len(all_trips)} 笔**  \n")
        f.write(f"**报销总金额 (根据行程单明细计算): {grand_total_detailed:.2f} 元**  \n\n")

        f.write("| 序号 | 日期 | 出发地 | 目的地 | 金额(元) |\n")
        f.write("|------|------|--------|--------|----------|\n")
        for i, trip in enumerate(all_trips, 1):
            f.write(f"| {i} | {trip.get('date', 'N/A')} | {trip.get('departure', 'N/A')} | {trip.get('destination', 'N/A')} | {trip.get('amount', 0.0):.2f} |\n")
        f.write("\n")

    # 同时生成发票号码汇总文件
    invoice_numbers = [res['invoice_number'] for res in comparison_results if res['invoice_number'] != "未找到发票"]
    invoice_numbers_filename = f'{directory_name}号码汇总.txt'
    with open(invoice_numbers_filename, 'w', encoding='utf-8') as f:
        f.write(",".join(invoice_numbers))

    print(f"\n汇总报告生成成功！已保存到: {output_filename}")
    print(f"发票号码汇总已保存到: {invoice_numbers_filename}")
    final_detailed_total = sum(t['amount'] for t in all_trips)
    print(f"摘要总金额: {grand_total_trip:.2f} | 明细计算总金额: {final_detailed_total:.2f}")
    if abs(grand_total_trip - final_detailed_total) < 0.01:
        print("所有金额完全匹配！任务圆满完成。")
    else:
        print(f"警告: 最终金额仍有差异: {abs(grand_total_trip - final_detailed_total):.2f}，请检查报告。")

if __name__ == "__main__":
    main()