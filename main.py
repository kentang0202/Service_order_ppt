import eel
import re
import datetime
import sys, os
from pptx import Presentation
from pptx.util import Pt

# 🏆 核心修復：精準抓取使用者執行檔所在的真實目錄
if getattr(sys, 'frozen', False):
    # 打包成 exe 執行時：指向 clickMeToRun.exe 所在的真實資料夾
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # 一般 python main.py 執行時：指向 main.py 所在資料夾
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 避免掃描子目錄造成 Google Drive 卡頓
_real_walk = os.walk
def _flat_walk(top, *args, **kwargs):
    for root, dirs, files in _real_walk(top, *args, **kwargs):
        dirs.clear()
        yield root, dirs, files

os.walk = _flat_walk
eel.init(BASE_DIR)
os.walk = _real_walk

def move_slide(prs, old_index, new_index):
    """將投影片從 old_index 移動到 new_index"""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[old_index])
    xml_slides.insert(new_index, slides[old_index])


def calculate_topic_font_size(topic_str: str) -> Pt:
    lines = [line.strip() for line in topic_str.split('\n') if line.strip()]
    if not lines:
        return Pt(100)
    
    max_chars = max(len(line) for line in lines)
    line_count = len(lines)
    
    # 🏆 調整為：單行字數 * 字級 ≈ 1100
    base_size = 1100 / max(max_chars, 1)
    
    # 行數扣減規則：4 行扣 10pt、5 行扣 20pt...
    if line_count >= 4:
        penalty = (line_count - 3) * 10
        base_size -= penalty
        
    # 上限 200pt、下限 80pt
    target_size = max(80, min(200, round(base_size)))
    
    return Pt(target_size)

def split_bible_text(text: str, max_chars: int = 100) -> list[str]:
    text = text.strip()
    # 1. 優先判斷：總字數未超過 100 字，直接保留原樣不拆
    if len(text) <= max_chars:
        return [text]

    # 2. 依自然換行符號拆分成段落
    raw_paragraphs = [p.strip() for p in text.split('\n') if p.strip()]
    
    # 3. 處理單一換行段落是否仍超過 100 字（若是，則用標點符號再細切）
    refined_units = []
    # 支援複合標點（如 。」 ！」 ？」）與一般標點（。！？；）
    punct_pattern = r'([。！？；][」』”"]?|[。！？；])'

    for p in raw_paragraphs:
        if len(p) <= max_chars:
            refined_units.append(p)
        else:
            # 超過 100 字的長段落，依標點符號拆解子句
            parts = re.split(punct_pattern, p)
            temp_sentence = ""
            for i in range(0, len(parts) - 1, 2):
                sentence = parts[i] + parts[i+1]
                if sentence.strip():
                    refined_units.append(sentence.strip())
            # 若末尾有未帶標點的文字，也補進去
            if len(parts) % 2 == 1 and parts[-1].strip():
                refined_units.append(parts[-1].strip())

    # 4. 貪婪合併：依序將各個單位合併進投影片，只要加起來不超過 100 字就放同一頁
    chunks = []
    current_chunk = ""

    for unit in refined_units:
        # 決定換行符號（若原本是獨立段落則換行保留格式）
        separator = "\n" if current_chunk else ""
        tentative = current_chunk + separator + unit

        if len(tentative) <= max_chars:
            current_chunk = tentative
        else:
            if current_chunk:
                chunks.append(current_chunk)
            
            # 若單一子句依然極端超過 100 字（沒有任何標點），強制切分
            while len(unit) > max_chars:
                chunks.append(unit[:max_chars])
                unit = unit[max_chars:]
            current_chunk = unit

    if current_chunk:
        chunks.append(current_chunk)

    return chunks


@eel.expose
def start_process(data):
    try:
        template_type = data.get('template_type', 'main_sunday')
        # 🏆 接收前端 select 的文字，若無則 fallback 到 cfg['label']
        template_display_name = data.get('template_name', '')

        # 模板組態設定：定義各自的底稿檔案與專屬的頁碼索引
        template_configs = {
            "main_sunday": {
                "file": "2026主日司會模板python.pptx",
                "label": "主日司會流程",
                "service_text": "主日禮拜",
                "has_proverbs": True,
                "idx_date_2": 8,       # 第二個日期頁面
                "idx_contents": 9,     # 經文目錄頁
                "idx_first_bible": 10, # 第一頁經文
                "idx_topic": 11        # 話語主題頁
            },
            "main_wednesday": {
                "file": "2026週三司會模板python.pptx",
                "label": "週三司會流程",
                "service_text": "週三禮拜",
                "has_proverbs": False,
                "idx_date_2": None, # 週三沒有這頁，設為 None
                "idx_contents": 2,
                "idx_first_bible": 3,
                "idx_topic": 4
            }
        }

        if template_type not in template_configs:
            return {
                "success": False,
                "status": f"❌ 尚未支援或未知的模板類型：{template_type}",
                "preview": ""
            }

        cfg = template_configs[template_type]
        template_path = os.path.join(BASE_DIR, cfg["file"])

        if not os.path.exists(template_path):
            return {
                "success": False,
                "status": f"❌ 找不到模板檔案：請確認 '{cfg['file']}' 是否放在程式旁。",
                "preview": ""
            }

        prs = Presentation(template_path)
        bible_raw = data.get('bible_raw', '').strip()
        proverbs_raw = data.get('proverbs_raw', '').strip()
        topic_str = data.get('topic', '').strip()

        # 1. 智慧型日期偵測
        date_match = re.search(r'(\d{4})[年./](\d{1,2})[月./](\d{1,2})', bible_raw)
        if date_match:
            y, m, d = date_match.groups()
            date_str = f"{y}年{int(m)}月{int(d)}日"
            file_date = f"{y}{int(m):0>2}{int(d):0>2}"
            date_display = f"📅 偵測日期：{date_str}"
        else:
            today = datetime.date.today()
            date_str = f"{today.year}年{today.month}月{today.day}日"
            file_date = today.strftime("%Y%m%d")
            date_display = f"📅 偵測日期：{date_str} <span class='auto-date'>(無輸入，使用當前日期)</span>"

        # 2. Regex 解析經文與箴言
        bible_matches = re.findall(r'〈(.*?)〉\s*(.*?)(?=\n?〈|$)', bible_raw, re.S)
        prov_regex = r'(司會|會眾|全體)\s*[–—－-]\s*(.*?)(?=\n?(?:司會|會眾|全體)\s*[–—－-]|(?:\(阿們！.*?\))?$|$)'
        proverbs_matches = re.findall(prov_regex, proverbs_raw, re.S)

        if not bible_matches:
            return {"success": False, "status": "❌ 辨識失敗：經文請確保包含 〈〉 括號標題。", "preview": ""}

        bible_titles = [m[0].strip() for m in bible_matches]

        # 3. 填寫共用日期與目錄
        prs.slides[1].shapes.title.text = date_str

        # 只有主日需要填寫第二個日期頁
        if cfg["idx_date_2"] is not None and cfg["idx_date_2"] < len(prs.slides):
            prs.slides[cfg["idx_date_2"]].shapes.title.text = date_str

        prs.slides[cfg["idx_contents"]].shapes.title.text = "\n".join(bible_titles)

        # 4. 箴言填寫（只有主日執行）
        if cfg["has_proverbs"]:
            for i, (role, text) in enumerate(proverbs_matches):
                if i < 5:
                    prs.slides[3 + i].shapes.title.text = text.strip()

        # --- 5. 經文填寫與動態加頁 (支援過長自動拆頁) ---
        first_bible_idx = cfg["idx_first_bible"]

        # A. 展開所有經文：超過 100 字自動拆頁，並標示 (1/3)、(2/3)...
        expanded_bibles = []
        for title, raw_content in bible_matches:
            chunks = split_bible_text(raw_content, max_chars=100)
            total_pages = len(chunks)

            if total_pages == 1:
                # 只有一頁時維持原標題，不加括號
                expanded_bibles.append((title, chunks[0]))
            else:
                # 多頁時標註 (頁碼/總頁數)
                for seq, chunk in enumerate(chunks, 1):
                    expanded_bibles.append((f"{title} ({seq}/{total_pages})", chunk))

        # B. 填寫底稿既有的第一頁經文
        first_title, first_content = expanded_bibles[0]
        prs.slides[first_bible_idx].shapes.title.text = first_title
        try:
            prs.slides[first_bible_idx].placeholders[1].text = first_content
        except:
            pass

        # C. 第 2 頁之後的經文：動態取得第一頁的母片版型進行 add_slide (避免 index 越界)
        bible_layout = prs.slides[first_bible_idx].slide_layout
        for title, content in expanded_bibles[1:]:
            new_slide = prs.slides.add_slide(bible_layout)
            new_slide.shapes.title.text = title
            try:
                new_slide.placeholders[1].text = content
            except:
                pass


        # 6. 主題頁填寫
        topic_slide = prs.slides[cfg["idx_topic"]]
        titles_joined = "、".join(bible_titles)

        # 動態計算最適字級
        topic_font_size = calculate_topic_font_size(topic_str)
        
        for shape in topic_slide.placeholders:
            if shape.placeholder_format.idx == 0:
                shape.text = topic_str
                # 遍歷段落並統一設定計算後的字級
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.font.size = topic_font_size
                    # 若需要強制置中對齊可加這行：
                    # from pptx.enum.text import PP_ALIGN
                    # paragraph.alignment = PP_ALIGN.CENTER
            elif shape.placeholder_format.idx == 1:
                shape.text = titles_joined
            elif shape.placeholder_format.idx == 21:
                shape.text = f"{date_str} {cfg['service_text']}"


        # 7. 新增回顧頁並將主題頁搬到最後一頁
        review_slide = prs.slides.add_slide(prs.slide_layouts[1])
        review_slide.shapes.title.text = date_str
        move_slide(prs, cfg["idx_topic"], len(prs.slides) - 1)

        # 8. 儲存檔案
        output_name = f"{file_date}{cfg['label']}.pptx"
        prs.save(os.path.join(BASE_DIR, output_name))

        # 9. 組裝前端預覽 HTML（採用模板改為 select 選項文字）
        display_template = template_display_name if template_display_name else cfg['label']
        preview_html = f"<div class='preview-header'>📋 採用模板：{display_template}</div>"
        preview_html += f"<div class='preview-header'>{date_display}</div>"
        formatted_topic = topic_str.strip().replace('\n', '<br>')
        preview_html += (
            f"<div class='preview-header'>"
            f"🖊️ 話語主題：<br>"
            f"<span class='topic-content'>{formatted_topic if formatted_topic else '（尚未輸入）'}</span>"
            f"</div><hr>"
        )

        preview_html += "<b>📖 辨識到的經文（含自動分頁）：</b><br>"
        for title, content in expanded_bibles:
            formatted_content = content.strip().replace('\n', '<br>')
            preview_html += (
                f"<div style='margin-bottom:12px; background:#fff; padding:10px; "
                f"border-radius:5px; border:1px solid #eee;'>"
                f"<b>【{title}】</b><br>{formatted_content}</div>"
            )

        if cfg["has_proverbs"]:
            preview_html += "<hr><b>💡 辨識到的箴言：</b><br>"
            for i, (role, text) in enumerate(proverbs_matches):
                preview_html += f"<div style='margin-bottom:8px;'><b>{i+1}. {role}：</b> {text.strip()}</div>"

        return {
            "success": True,
            "status": f"✅ 成功生成：{output_name}",
            "preview": preview_html
        }

    except Exception as e:
        return {
            "success": False,
            "status": f"❌ 系統錯誤：{str(e)}",
            "preview": ""
        }

eel.start('index.html', size=(900, 800))