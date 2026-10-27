# -*- coding: utf-8 -*-
import zipfile
import os
import html

def escape_xml(s):
    return html.escape(str(s))

def make_docx(output_path):
    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Leelawadee UI" w:hAnsi="Leelawadee UI" w:cs="TH Sarabun New"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="28"/>
        <w:color w:val="2D2422"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="320" w:lineRule="auto" w:after="160"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>"""

    # Helper functions to build Word XML
    body_xml_parts = []

    def p_title(text):
        return f"""<w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="300" w:after="100"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:b/>
              <w:sz w:val="48"/>
              <w:szCs w:val="56"/>
              <w:color w:val="A66C44"/>
            </w:rPr>
            <w:t>{escape_xml(text)}</w:t>
          </w:r>
        </w:p>"""

    def p_subtitle(text):
        return f"""<w:p>
          <w:pPr>
            <w:jc w:val="center"/>
            <w:spacing w:before="0" w:after="400"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:sz w:val="28"/>
              <w:szCs w:val="32"/>
              <w:color w:val="654321"/>
            </w:rPr>
            <w:t>{escape_xml(text)}</w:t>
          </w:r>
        </w:p>"""

    def p_h1(text):
        return f"""<w:p>
          <w:pPr>
            <w:spacing w:before="400" w:after="160"/>
            <w:pBdr>
              <w:bottom w:val="single" w:sz="12" w:space="4" w:color="A66C44"/>
            </w:pBdr>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:b/>
              <w:sz w:val="36"/>
              <w:szCs w:val="40"/>
              <w:color w:val="4A3525"/>
            </w:rPr>
            <w:t>{escape_xml(text)}</w:t>
          </w:r>
        </w:p>"""

    def p_h2(text):
        return f"""<w:p>
          <w:pPr>
            <w:spacing w:before="240" w:after="120"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:b/>
              <w:sz w:val="30"/>
              <w:szCs w:val="34"/>
              <w:color w:val="A66C44"/>
            </w:rPr>
            <w:t>{escape_xml(text)}</w:t>
          </w:r>
        </w:p>"""

    def p_h3(text):
        return f"""<w:p>
          <w:pPr>
            <w:spacing w:before="160" w:after="80"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:b/>
              <w:sz w:val="26"/>
              <w:szCs w:val="30"/>
              <w:color w:val="2D2422"/>
            </w:rPr>
            <w:t>{escape_xml(text)}</w:t>
          </w:r>
        </w:p>"""

    def p_para(text, bold_prefix=""):
        res = "<w:p><w:pPr><w:spacing w:after=\"120\"/></w:pPr>"
        if bold_prefix:
            res += f"""<w:r>
              <w:rPr><w:b/><w:color w:val="4A3525"/></w:rPr>
              <w:t>{escape_xml(bold_prefix)} </w:t>
            </w:r>"""
        res += f"""<w:r>
          <w:t>{escape_xml(text)}</w:t>
        </w:r></w:p>"""
        return res

    def p_bullet(text, bold_title=""):
        res = """<w:p>
          <w:pPr>
            <w:ind w:left="400" w:hanging="200"/>
            <w:spacing w:after="80"/>
          </w:pPr>
          <w:r>
            <w:rPr><w:color w:val="A66C44"/></w:rPr>
            <w:t>• </w:t>
          </w:r>"""
        if bold_title:
            res += f"""<w:r>
              <w:rPr><w:b/><w:color w:val="3B2A1E"/></w:rPr>
              <w:t>{escape_xml(bold_title)}: </w:t>
            </w:r>"""
        res += f"""<w:r>
          <w:t>{escape_xml(text)}</w:t>
        </w:r></w:p>"""
        return res

    def p_code_box(code_text):
        return f"""<w:p>
          <w:pPr>
            <w:spacing w:before="100" w:after="160"/>
            <w:shd w:val="clear" w:color="auto" w:fill="F5EFE6"/>
            <w:ind w:left="300" w:right="300"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>
              <w:b/>
              <w:sz w:val="22"/>
              <w:szCs w:val="24"/>
              <w:color w:val="8C5832"/>
            </w:rPr>
            <w:t>{escape_xml(code_text)}</w:t>
          </w:r>
        </w:p>"""

    def p_note_box(note_text, label="ข้อแนะนำ"):
        return f"""<w:p>
          <w:pPr>
            <w:spacing w:before="160" w:after="160"/>
            <w:shd w:val="clear" w:color="auto" w:fill="FAF0E6"/>
            <w:pBdr>
              <w:left w:val="single" w:sz="24" w:space="12" w:color="A66C44"/>
            </w:pBdr>
            <w:ind w:left="300" w:right="200"/>
          </w:pPr>
          <w:r>
            <w:rPr><w:b/><w:color w:val="A66C44"/></w:rPr>
            <w:t>💡 {escape_xml(label)}: </w:t>
          </w:r>
          <w:r>
            <w:t>{escape_xml(note_text)}</w:t>
          </w:r>
        </w:p>"""

    def make_table(headers, rows, col_widths=None):
        num_cols = len(headers)
        if not col_widths:
            w = 9000 // num_cols
            col_widths = [w] * num_cols

        tbl = """<w:tbl>
          <w:tblPr>
            <w:tblW w:w="9000" w:type="dxa"/>
            <w:jc w:val="center"/>
            <w:tblBorders>
              <w:top w:val="single" w:sz="4" w:space="0" w:color="E8DEC8"/>
              <w:left w:val="none"/>
              <w:bottom w:val="single" w:sz="12" w:space="0" w:color="A66C44"/>
              <w:right w:val="none"/>
              <w:insideH w:val="single" w:sz="4" w:space="0" w:color="F0E6D8"/>
              <w:insideV w:val="none"/>
            </w:tblBorders>
          </w:tblPr>"""

        # Header row
        tbl += "<w:tr><w:trPr><w:tblHeader/></w:trPr>"
        for i, h in enumerate(headers):
            tbl += f"""<w:tc>
              <w:tcPr>
                <w:tcW w:w="{col_widths[i]}" w:type="dxa"/>
                <w:shd w:val="clear" w:color="auto" w:fill="FAF0E6"/>
                <w:tcMar>
                  <w:top w:w="120" w:type="dxa"/>
                  <w:bottom w:w="120" w:type="dxa"/>
                  <w:left w:w="140" w:type="dxa"/>
                  <w:right w:w="140" w:type="dxa"/>
                </w:tcMar>
              </w:tcPr>
              <w:p>
                <w:pPr><w:spacing w:after="0"/></w:pPr>
                <w:r><w:rPr><w:b/><w:color w:val="4A3525"/></w:rPr><w:t>{escape_xml(h)}</w:t></w:r>
              </w:p>
            </w:tc>"""
        tbl += "</w:tr>"

        # Body rows
        for row in rows:
            tbl += "<w:tr>"
            for i, cell in enumerate(row):
                tbl += f"""<w:tc>
                  <w:tcPr>
                    <w:tcW w:w="{col_widths[i]}" w:type="dxa"/>
                    <w:tcMar>
                      <w:top w:w="100" w:type="dxa"/>
                      <w:bottom w:w="100" w:type="dxa"/>
                      <w:left w:w="140" w:type="dxa"/>
                      <w:right w:w="140" w:type="dxa"/>
                    </w:tcMar>
                  </w:tcPr>
                  <w:p>
                    <w:pPr><w:spacing w:after="0"/></w:pPr>
                    <w:r><w:t>{escape_xml(cell)}</w:t></w:r>
                  </w:p>
                </w:tc>"""
            tbl += "</w:tr>"

        tbl += "</w:tbl><w:p><w:pPr><w:spacing w:after=\"160\"/></w:pPr></w:p>"
        return tbl

    # Assemble Document Content
    # Title & Subtitle
    body_xml_parts.append(p_title("คู่มือการใช้งานระบบเว็บเมนูร้านอาหาร & สั่งซื้อ"))
    body_xml_parts.append(p_subtitle("Komorebi & Co. - East Meets West Artisan Patisserie & Cafe"))

    # Overview
    body_xml_parts.append(p_note_box(
        "ระบบเว็บเมนูและสั่งซื้ออาหาร Komorebi & Co. พัฒนาขึ้นด้วยเทคโนโลยีโมเดิร์น React 19, TypeScript และ Tailwind CSS "
        "รองรับการใช้งานได้ทั้งลูกค้าหน้าร้าน (สแกนดูเมนูผ่านมือถือ, เลือกโต๊ะ, ปรับแต่งความหวาน/ท็อปปิ้ง, สแกนจ่ายเงิน QR Code) "
        "และพนักงานในร้าน (ระบบ POS หลังบ้าน ติดตามออเดอร์ในครัวแบบเรียลไทม์ และควบคุมสต็อกสินค้า)",
        "ภาพรวมระบบ (System Overview)"
    ))

    # Chapter 1
    body_xml_parts.append(p_h1("บทที่ 1: การเปิดใช้งานระบบ (Getting Started)"))
    body_xml_parts.append(p_para("ระบบถูกติดตั้งและตั้งค่าเสร็จสมบูรณ์ในโฟลเดอร์ของเครื่องเรียบร้อยแล้ว โดยสามารถเปิดใช้งานได้ตามขั้นตอนดังนี้:"))
    body_xml_parts.append(p_h2("1.1 การรันระบบผ่าน Terminal / Command Prompt"))
    body_xml_parts.append(p_para("เปิด Command Prompt หรือ PowerShell ในโฟลเดอร์ C:\\Pornprom แล้วพิมพ์คำสั่ง:"))
    body_xml_parts.append(p_code_box("npm run dev"))
    body_xml_parts.append(p_para("หรือสามารถใช้คำสั่งตรงของ Vite ได้เช่นกัน:"))
    body_xml_parts.append(p_code_box("npx vite"))
    body_xml_parts.append(p_h2("1.2 การเข้าใช้งานผ่านเว็บเบราว์เซอร์"))
    body_xml_parts.append(p_para("เมื่อเซิร์ฟเวอร์เริ่มทำงาน ให้เปิดเบราว์เซอร์ (Google Chrome, Microsoft Edge, Safari) แล้วไปที่ URL:"))
    body_xml_parts.append(p_code_box("http://localhost:5173"))
    body_xml_parts.append(p_bullet("สามารถเปิดดูได้ทั้งบนหน้าจอคอมพิวเตอร์ แท็บเล็ต และสมาร์ตโฟน โดยหน้าจอจะปรับเลย์เอาต์อัตโนมัติ (Responsive Design)", "ความยืดหยุ่น"))

    # Chapter 2
    body_xml_parts.append(p_h1("บทที่ 2: ฟังก์ชันสำหรับลูกค้าหน้าร้าน (Customer Menu & Ordering)"))
    body_xml_parts.append(p_para("หน้าเว็บส่วนของลูกค้าถูกออกแบบให้อ่านง่าย รูปภาพคมชัด และสั่งซื้อได้ง่ายโดยไม่ต้องโหลดแอปพลิเคชัน:"))
    
    body_xml_parts.append(p_h2("2.1 การเลือกโหมดการสั่งซื้อ (Order Mode)"))
    body_xml_parts.append(p_para("ลูกค้าสามารถเลือกโหมดบริการได้จากแถบเมนูด้านบน:"))
    body_xml_parts.append(p_bullet("สำหรับลูกค้านั่งทานในร้าน สามารถระบุหมายเลขโต๊ะในขั้นตอนชำระเงินเพื่อให้อาหารและเครื่องดื่มไปเสิร์ฟที่โต๊ะได้ถูกต้อง", "1. สั่งทานตอนนี้ / หน้าร้าน (Dine-in)"))
    body_xml_parts.append(p_bullet("สำหรับสั่งนำกลับบ้าน เมื่อชำระเงินจะได้รับหมายเลขคิวเพื่อรอรับที่เคาน์เตอร์ Grab & Go", "2. สั่งกลับบ้าน (Takeaway)"))
    body_xml_parts.append(p_bullet("สำหรับสั่งจองเค้กปอนด์หรือ Boxset ของขวัญ โดยสามารถเลือกวันที่และเวลานัดรับล่วงหน้าได้", "3. สั่งจองล่วงหน้า (Pre-order)"))

    body_xml_parts.append(p_h2("2.2 การค้นหาและกรองหมวดหมู่เมนู"))
    body_xml_parts.append(p_bullet("คลิกเลือกแท็บเพื่อดูเฉพาะเมนูประเภทที่ต้องการ เช่น 'ครัวซองต์ & พัฟ', 'เค้ก & ทาร์ต', 'ขนมปัง & สโคน', 'กาแฟสด', 'ชา & รีเฟรชเชอร์', หรือ 'เซ็ตของขวัญ'", "แถบหมวดหมู่ (Category Pills)"))
    body_xml_parts.append(p_bullet("พิมพ์คำค้นหา เช่น 'มัทฉะ', 'ยูซุ', 'ลาวา' ระบบจะแสดงผลลัพธ์ทันทีแบบ Real-time", "ช่องค้นหาด่วน (Search Box)"))

    body_xml_parts.append(p_h2("2.3 การปรับแต่งรายละเอียดเมนู (Customize Modal)"))
    body_xml_parts.append(p_para("เมื่อคลิกที่การ์ดเมนูใดๆ หน้าต่างป๊อปอัปจะเปิดขึ้นเพื่อให้ลูกค้าเลือกความต้องการพิเศษ:"))
    body_xml_parts.append(p_bullet("เลือกระดับความหวาน (0% ไม่หวาน, 25% หวานน้อยมาก, 50% หวานน้อยกำลังดี, 100% หวานปกติ), ประเภทนม (นมสดปกติ, นมโอ๊ต +20 บาท, นมอัลมอนด์ +20 บาท), รูปแบบอุณหภูมิ (ร้อน / เย็น / ปั่น +15 บาท), เพิ่มช็อตกาแฟ หรือท็อปปิ้งโฟมชีส", "หมวดเครื่องดื่ม (Coffee & Tea)"))
    body_xml_parts.append(p_bullet("เลือกให้อุ่นร้อนก่อนเสิร์ฟได้ฟรี (ทางร้านจะนำเข้าเตาอบให้กรอบหอมก่อนเสิร์ฟ)", "หมวดเบเกอรี่ (Bakery)"))
    body_xml_parts.append(p_bullet("เลือกขนาดเค้ก (ชิ้นเดี่ยว, 1 ปอนด์, 2 ปอนด์), เขียนข้อความอวยพรหน้าเค้ก และเลือกจำนวนเทียนวันเกิด", "หมวดเค้ก (Cakes)"))

    body_xml_parts.append(p_h2("2.4 ตะกร้าสินค้าและการชำระเงิน (Cart & Checkout Flow)"))
    body_xml_parts.append(p_bullet("กดปุ่ม 'ตะกร้า' ด้านบนหรือแถบแจ้งเตือนด้านล่างเพื่อตรวจสอบรายการสินค้า ปรับเพิ่ม-ลดจำนวน หรือลบรายการ", "1. ตรวจสอบตะกร้า"))
    body_xml_parts.append(p_bullet("กรอกชื่อผู้สั่ง และเบอร์โทรศัพท์ (และระบุเลขโต๊ะ หรือวัน-เวลานัดรับ)", "2. ระบุข้อมูลผู้สั่ง"))
    body_xml_parts.append(p_bullet("ระบบจะคำนวณยอดเงินรวมสุทธิ และแสดง QR Code พร้อมเพย์ (PromptPay) พร้อมปุ่มคัดลอกเลขบัญชี", "3. ชำระเงินด้วย QR Code"))
    body_xml_parts.append(p_bullet("เมื่อชำระเงินสำเร็จ ระบบจะแสดงหน้าต่างขอบคุณพร้อมหมายเลขคิว (Queue No. เช่น K-01, K-02) เพื่อให้ลูกค้านำไปแสดงตอนรับอาหาร", "4. รับหมายเลขคิว (Queue Code)"))

    # Chapter 3
    body_xml_parts.append(p_h1("บทที่ 3: ระบบจัดการร้านค้าและครัว (Staff POS & Kitchen Board)"))
    body_xml_parts.append(p_para("ร้านค้าสามารถสลับเข้าสู่โหมดพนักงานได้โดยคลิกปุ่ม **'ระบบร้านค้า (POS/ครัว)'** ที่มุมขวาบนของแถบนำทาง (Navbar) เพื่อจัดการงานภายในร้านได้อย่างสะดวกรวดเร็ว:"))

    body_xml_parts.append(p_h2("3.1 แดชบอร์ดสรุปยอดขายด่วน (Realtime KPI)"))
    body_xml_parts.append(p_bullet("แสดงจำนวนออเดอร์ที่ลูกค้าเพิ่งสั่งเข้ามาและกำลังรอการจัดเตรียมในครัว", "ออเดอร์ที่กำลังรอทำ (Pending & Preparing)"))
    body_xml_parts.append(p_bullet("แสดงจำนวนออเดอร์ที่ทำเสร็จแล้ว กำลังรอพนักงานนำไปเสิร์ฟที่โต๊ะ หรือรอลูกค้ามารับที่เคาน์เตอร์", "พร้อมเสิร์ฟ / รอรับ (Ready)"))
    body_xml_parts.append(p_bullet("คำนวณยอดเงินรวมจากออเดอร์ที่ชำระเงินแล้วในแต่ละวันแบบสดๆ", "ยอดขายรวมวันนี้ (Total Revenue)"))

    body_xml_parts.append(p_h2("3.2 กระดานติดตามออเดอร์ (Live Orders Board)"))
    body_xml_parts.append(p_para("พนักงานและบาริสตาสามารถเห็นออเดอร์ที่เข้ามา พร้อมรายละเอียดการปรับแต่ง (เช่น ระดับความหวาน, ให้อุ่นร้อน, ข้อความหน้าเค้ก) โดยสามารถกดเปลี่ยนสถานะได้ด้วยคลิกเดียว:"))
    body_xml_parts.append(p_bullet("เมื่อลูกค้าสั่งเข้ามาใหม่ ออเดอร์จะขึ้นแท็กสีเหลือง 'รอดำเนินการ'", "สถานะที่ 1: รอดำเนินการ (Pending)"))
    body_xml_parts.append(p_bullet("กดปุ่ม 'เริ่มทำออเดอร์' แท็กจะเปลี่ยนเป็นสีน้ำเงิน เพื่อให้เพื่อนร่วมงานทราบว่ากำลังทำอยู่", "สถานะที่ 2: กำลังจัดเตรียม (Preparing)"))
    body_xml_parts.append(p_bullet("เมื่อเบเกอรี่อบเสร็จหรือชงเครื่องดื่มเสร็จ กดปุ่ม 'ทำเสร็จแล้ว' แท็กจะเปลี่ยนเป็นสีเขียว เพื่อเตรียมเสิร์ฟ", "สถานะที่ 3: พร้อมเสิร์ฟ (Ready)"))
    body_xml_parts.append(p_bullet("เมื่อนำอาหารไปเสิร์ฟที่โต๊ะหรือลูกค้ามารับเรียบร้อย กดปุ่ม 'เสิร์ฟเรียบร้อย' เพื่อปิดออเดอร์", "สถานะที่ 4: เสร็จสิ้น (Completed)"))

    body_xml_parts.append(p_h2("3.3 ระบบคุมสต็อกสินค้าหน้าร้าน (Inventory & Stock Manager)"))
    body_xml_parts.append(p_para("ในแท็บ 'คุมสต็อกสินค้า' พนักงานสามารถบริหารจัดการขนมและเครื่องดื่มได้อย่างยืดหยุ่น:"))
    body_xml_parts.append(p_bullet("หากเบเกอรี่รายการใดอบหมดแล้ว สามารถกดปุ่มเพื่อสลับสถานะเป็น 'หมดชั่วคราว' ป้ายบนหน้าเว็บลูกค้าจะขึ้นว่าสินค้าหมดทันที และไม่สามารถกดสั่งได้", "1. สลับสถานะของหมด (Toggle Availability)"))
    body_xml_parts.append(p_bullet("สามารถกดปุ่ม [+] หรือ [-] เพื่ออัปเดตจำนวนชิ้นของครัวซองต์หรือเค้กที่อบเสร็จใหม่ในแต่ละรอบได้ตลอดเวลา", "2. ปรับจำนวนสต็อก (Adjust Stock Count)"))
    body_xml_parts.append(p_bullet("เมื่อลูกค้าสั่งซื้อขนม ระบบจะตัดสต็อกคงเหลือให้อัตโนมัติ", "3. ตัดสต็อกอัตโนมัติ (Auto Deduct)"))

    # Chapter 4
    body_xml_parts.append(p_h1("บทที่ 4: รายการเมนูและราคามาตรฐาน (Menu Catalog & Pricing)"))
    body_xml_parts.append(p_para("ตารางรายการเมนูมาตรฐานที่ติดตั้งไว้ในระบบ พร้อมราคาจำหน่าย:"))

    body_xml_parts.append(p_h2("4.1 รายการเบเกอรี่และของหวาน (Bakery & Pastries)"))
    bakery_headers = ["รหัส", "ชื่อเมนูภาษาไทย", "ชื่อเมนูภาษาอังกฤษ", "ประเภท", "ราคา"]
    bakery_rows = [
        ["p0", "ครัวซองต์เนยสดฝรั่งเศส", "Classic French Butter Croissant", "Viennoiserie", "95 บาท"],
        ["p1", "ครัวซองต์สังขยาใบเตยมะพร้าวอ่อน", "Pandan & Young Coconut Croissant", "Signature", "145 บาท"],
        ["p2", "ครัวซองต์มัทฉะโมจิยืด", "Uji Matcha & Mochi Croissant", "Japanese Twist", "165 บาท"],
        ["p3", "ครัวซองต์ไข่เค็มลาวาไชยา", "Chaiya Salted Egg Lava Croissant", "Thai Twist", "155 บาท"],
        ["p4", "ชูครีมกรอบงาดำคั่วบด", "Black Sesame Craquelin Choux", "Choux Cream", "120 บาท"],
        ["c1", "ทาร์ตส้มยูซุและเมอแรงก์อิตาเลียน", "Yuzu Lemon Meringue Tart", "Tart", "185 บาท"],
        ["c2", "ชีสเค้กหน้าไหม้รสชาไทยเข้มข้น", "Thai Tea Basque Burnt Cheesecake", "Signature Cake", "175 บาท"],
        ["c3", "เค้กมะพร้าวอ่อนนมฮอกไกโด (ปอนด์)", "Young Coconut Hokkaido Shortcake", "Pre-order 1 วัน", "890 บาท"],
        ["b1", "สโคนชาเอิร์ลเกรย์ & บราวชูการ์โบบา", "Earl Grey & Brown Sugar Boba Scone", "Scone", "135 บาท"],
        ["b2", "ขนมปังบริออชสับปะรดภูแลฉ่ำเนย", "Phulae Pineapple Brioche", "Artisan Bread", "125 บาท"],
    ]
    body_xml_parts.append(make_table(bakery_headers, bakery_rows, [800, 2800, 2700, 1500, 1200]))

    body_xml_parts.append(p_h2("4.2 รายการเครื่องดื่มคราฟต์ (Coffee, Tea & Refreshers)"))
    drink_headers = ["รหัส", "ชื่อเมนูภาษาไทย", "ชื่อเมนูภาษาอังกฤษ", "ประเภท", "ราคา"]
    drink_rows = [
        ["cf1", "ยูซุ เอสเปรสโซ โทนิค", "Yuzu Espresso Tonic", "Specialty Coffee", "140 บาท"],
        ["cf2", "เดอร์ตี้ ชาไทยเอสเปรสโซ", "Dirty Thai Tea Espresso", "Signature Fusion", "135 บาท"],
        ["cf3", "โคโคนัท คลาวด์ อเมริกาโน่", "Coconut Cloud Americano", "Specialty Coffee", "130 บาท"],
        ["cf4", "ดาร์กช็อกโกแลตเย็นพรีเมียม 70%", "Craft Dark Chocolate 70%", "Single Origin", "130 บาท"],
        ["t1", "มัทฉะลาเต้ อูจิเกรดพิธีการ", "Ceremonial Uji Matcha Latte", "Specialty Tea", "145 บาท"],
        ["t2", "ชาอู่หลงลิ้นจี่ทับทิมสปาร์คกลิ้ง", "Lychee Ruby Oolong Sparkling", "Refresher", "125 บาท"],
        ["t3", "โฮจิฉะลาเต้คั่วหอม", "Roasted Hojicha Latte", "Specialty Tea", "135 บาท"],
        ["g1", "East Meets West Boxset (4 ชิ้น)", "Artisan Signature Giftset Box", "Gift Set", "560 บาท"],
    ]
    body_xml_parts.append(make_table(drink_headers, drink_rows, [800, 2800, 2700, 1500, 1200]))

    # Chapter 5
    body_xml_parts.append(p_h1("บทที่ 5: การปรับแต่งและแก้ไขข้อมูลระบบ (System Customization)"))
    body_xml_parts.append(p_para("สำหรับเจ้าของร้านหรือผู้ดูแลระบบที่ต้องการแก้ไขข้อมูล สามารถดำเนินการได้ง่ายๆ ดังนี้:"))
    body_xml_parts.append(p_h2("5.1 การเพิ่มหรือแก้ไขเมนูและราคา"))
    body_xml_parts.append(p_para("ข้อมูลเมนูอาหารทั้งหมดจัดเก็บอยู่ที่ไฟล์:"))
    body_xml_parts.append(p_code_box("src/data/mockData.ts"))
    body_xml_parts.append(p_bullet("สามารถเปิดไฟล์นี้ด้วยโปรแกรม Text Editor (เช่น VS Code, Notepad) เพื่อเปลี่ยนชื่อเมนู, แก้ไขราคา (price), คำอธิบาย (description) หรือเปลี่ยนลิงก์รูปภาพ (imageUrl) ได้ทันที", "การแก้ไขข้อมูล"))

    body_xml_parts.append(p_h2("5.2 การปรับเปลี่ยนโทนสีและธีมร้าน"))
    body_xml_parts.append(p_para("สไตล์และการจัดวางถูกควบคุมด้วย Tailwind CSS ในไฟล์:"))
    body_xml_parts.append(p_code_box("src/index.css และ src/App.tsx"))
    body_xml_parts.append(p_bullet("โทนสีหลักปัจจุบันคือ สีครีมอบอุ่น (#FAF7F2), สีกาแฟและคาราเมล (#A66C44), และสีน้ำตาลเข้ม (#3B2A1E) ซึ่งเหมาะกับบรรยากาศร้านคาเฟ่และเบเกอรี่", "ธีมสี"))

    # Chapter 6
    body_xml_parts.append(p_h1("บทที่ 6: คำถามที่พบบ่อย (FAQ & Troubleshooting)"))
    body_xml_parts.append(p_bullet("ให้ตรวจสอบว่าได้เปิดโปรแกรมรันเซิร์ฟเวอร์ (npm run dev หรือ npx vite) ค้างไว้ใน Terminal หรือไม่ และเช็กว่าพอร์ต 5173 ไม่ถูกโปรแกรมอื่นใช้งานอยู่", "1. เปิดหน้าเว็บ http://localhost:5173 ไม่ได้"))
    body_xml_parts.append(p_bullet("หากต้องการนำขึ้นระบบออนไลน์ สามารถ Deploy ขึ้นแพลตฟอร์มฟรี เช่น Vercel, Netlify, หรือ Firebase Hosting ได้ง่ายๆ โดยใช้คำสั่ง npm run build แล้วอัปโหลดโฟลเดอร์ dist/", "2. ต้องการเปิดให้ลูกค้าภายนอกสแกน QR Code จริงบนมือถือ"))
    body_xml_parts.append(p_bullet("การเปลี่ยนสถานะหรือตัดสต็อกในหน้านี้เป็นการทำงานแบบ Realtime State ในเบราว์เซอร์ หากต้องการเชื่อมต่อฐานข้อมูลจริง เช่น Supabase หรือ Firebase สามารถเชื่อมต่อ API ต่อเติมได้ทันที", "3. ข้อมูลออเดอร์ในหน้า Staff POS จัดเก็บอย่างไร"))

    # Final wrap-up
    body_xml_parts.append(p_para("จัดทำโดย: ทีมงานพัฒนาระบบ Komorebi & Co. Artisan Patisserie & Cafe", "เอกสารเวอร์ชัน 1.0 (ตุลาคม 2026) |"))

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {"".join(body_xml_parts)}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    # Create ZIP / docx
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types_xml)
        zf.writestr('_rels/.rels', rels_xml)
        zf.writestr('word/_rels/document.xml.rels', doc_rels_xml)
        zf.writestr('word/styles.xml', styles_xml)
        zf.writestr('word/document.xml', document_xml)

    print(f"Successfully generated DOCX manual at: {output_path}")

if __name__ == "__main__":
    out_dir = r"C:\Pornprom"
    out_file = os.path.join(out_dir, "User_Manual_Komorebi_Cafe.docx")
    make_docx(out_file)
