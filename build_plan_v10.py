# -*- coding: utf-8 -*-
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/user/Feypk/Plan_RU_Econ_Quant_v10.xlsx"

FILL = {"econ":PatternFill("solid",fgColor="DDEBF7"),"quant":PatternFill("solid",fgColor="E2EFDA"),
        "fin":PatternFill("solid",fgColor="FCE4D6"),"gen":PatternFill("solid",fgColor="F2F2F2")}
HEADER_FILL=PatternFill("solid",fgColor="4472C4"); TERM_FILL=PatternFill("solid",fgColor="D9E1F2")
TITLE_FILL=PatternFill("solid",fgColor="1F4E78")
D_RED=PatternFill("solid",fgColor="FFC7CE"); D_ORG=PatternFill("solid",fgColor="F4B084")
D_YEL=PatternFill("solid",fgColor="FFE699"); D_GRN=PatternFill("solid",fgColor="C6E0B4")
def dfill(d): return D_RED if d>=9 else D_ORG if d>=7 else D_YEL if d>=5 else D_GRN
FONT=Font(name="Tahoma",size=10); FONT_B=Font(name="Tahoma",size=10,bold=True)
FONT_HDR=Font(name="Tahoma",size=10,bold=True,color="FFFFFF"); FONT_TITLE=Font(name="Tahoma",size=13,bold=True,color="FFFFFF")
FONT_NOTE=Font(name="Tahoma",size=9,italic=True)
thin=Side(style="thin",color="BFBFBF"); BORDER=Border(left=thin,right=thin,top=thin,bottom=thin)
CENTER=Alignment(horizontal="center",vertical="center"); LEFT=Alignment(horizontal="left",vertical="center",wrap_text=True)
HEADERS=["✅","รหัส","ชื่อวิชา","นก.","ยาก/10","หมวด / เหตุผล","บุพวิชา","สถานะ / หมายเหตุ"]

# row = (code, name, credits, diff, reason, prereq, status, cat)
PLAN=[
 ("ปีที่ 1 / เทอม 1",21,[
  ("ECO1121","หลักเศรษฐศาสตร์จุลภาค",3,3,"แกน econ","-","✓","econ"),
  ("ECO1123","คณิต–สถิติเบื้องต้นสำหรับ ศ.",3,4,"แกน • ด่านทดสอบเลข","-","✓","econ"),
  ("ACC1103","การบัญชีการเงิน",3,4,"แกน econ","-","✓","econ"),
  ("MTH1101","Calculus & Analytic Geometry I",3,4,"เลือกเสรี","-","✓","quant"),
  ("RAM1111","English in Daily Life",3,2,"ศึกษาทั่วไป 1.1","-","✓","gen"),
  ("RAM1103","Thai for Communication at Work",3,2,"ศึกษาทั่วไป 1.1","-","✓","gen"),
  ("COS1101","Introduction to Computer Science",3,3,"เก็บเพิ่ม • พื้นฐาน CS","-","✓","quant"),
 ]),
 ("ปีที่ 1 / เทอม 2",21,[
  ("ECO1122","หลักเศรษฐศาสตร์มหภาค",3,3,"แกน econ","ECO1123","✓","econ"),
  ("MTH1102","Calculus & Analytic Geometry II",3,5,"เลือกเสรี","MTH1101","✓","quant"),
  ("MTH2001","Fundamental Concepts in Math 1",3,7,"★ พิสูจน์ (ฐาน Real Analysis)","-","✓","quant"),
  ("MTH2206","Matrix Theory & Linear Algebra 1",3,5,"โทคณิต","MTH1101","✓","quant"),
  ("RAM1131","Digital Literacy",3,2,"ศึกษาทั่วไป 1.1","-","✓","gen"),
  ("RAM1301","คุณธรรมคู่ความรู้",3,2,"ศึกษาทั่วไป 1.3","-","✓","gen"),
  ("COS1103","Algorithms and Programming",3,5,"เก็บเพิ่ม • ปลดล็อก COS2102","COS1101","✓","quant"),
 ]),
 ("ปีที่ 2 / เทอม 1",21,[
  ("ECO2121","ทฤษฎีจุลภาค 1",3,6,"แกน econ","ECO1121,1123","✓","econ"),
  ("ECO2122","ทฤษฎีมหภาค 1",3,6,"แกน econ","ECO1122,1123","✓","econ"),
  ("ECO2125","คณิตศาสตร์สำหรับเศรษฐศาสตร์",3,6,"แกน • วิชาสัญญาณ","ECO1123","✓","econ"),
  ("ECO2129","เศรษฐกิจประเทศไทย",3,3,"แกน econ","ECO1121,1122","✓","econ"),
  ("MTH2103","Calculus & Analytic Geometry III",3,6,"โทคณิต • หลายตัวแปร","MTH1102","✓","quant"),
  ("STA3101","Probability Theory 1",3,7,"★ ฐานความน่าจะเป็น","MTH1102","✓","quant"),
  ("RAM1112","ภาษาและวัฒนธรรมอังกฤษ",3,2,"ศึกษาทั่วไป 1.1","-","✓","gen"),
 ]),
 ("ปีที่ 2 / เทอม 2",21,[
  ("ECO2123","ทฤษฎีจุลภาค 2",3,7,"แกน econ","ECO2121","✓","econ"),
  ("ECO2124","ทฤษฎีมหภาค 2",3,7,"แกน econ","ECO2122","✓","econ"),
  ("ECO2126","สถิติสำหรับเศรษฐศาสตร์",3,6,"แกน • วิชาสัญญาณ","STA2003 (ทางการ)","⚠️ ไม่มี STA2003 — เช็ก ECO1123 แทน","econ"),
  ("MTH2104","Advanced Calculus 1",3,8,"โทคณิต • ปลดล็อก Real Analysis","MTH2103","✓","quant"),
  ("MTH3204","Matrix Theory & Linear Algebra 2",3,7,"โทคณิต","MTH2206","✓","quant"),
  ("RAM1141","ศาสตร์แห่งบุคลิกภาพ",3,2,"ศึกษาทั่วไป 1.1","-","✓","gen"),
  ("ACC2202","Intermediate Accounting",3,6,"บัญชีเชิงลึก (ก่อนฝึกงาน)","ACC1103","⚠️ เช็ก ACC1101/1103","fin"),
 ]),
 ("ปีที่ 3 / เทอม 1",21,[
  ("ECO3121","เศรษฐศาสตร์จุลภาคระดับสูง",3,8,"เชี่ยวชาญ","ECO2123,2125","✓","econ"),
  ("ECO3123","เศรษฐมิติเบื้องต้น",3,7,"เชี่ยวชาญ • ดาวเด่น","ECO2125,2126","✓ (รอ ECO2126)","econ"),
  ("ECO3320","เศรษฐศาสตร์การเงินและการธนาคาร",3,5,"เลือกกลุ่มอื่น–การเงิน","ECO2124","✓","fin"),
  ("MTH2106","Differential Equations",3,6,"โทคณิต","MTH1102","✓","quant"),
  ("STA3102","Probability Theory 2",3,8,"★ ปลดล็อก Theory of Stats","STA3101+MTH2103","✓","quant"),
  ("ECO2130","English for Economics",3,3,"แกน econ","-","✓","econ"),
  ("FIN2101","Business Finance",3,5,"เก็บเพิ่ม • บุพวิชา FIN3xxx","ACC1103","⚠️ เช็ก PR + สิทธิ์ข้ามคณะ","fin"),
 ]),
 ("ปีที่ 3 / เทอม 2",21,[
  ("ECO3122","เศรษฐศาสตร์มหภาคระดับสูง",3,8,"เชี่ยวชาญ","ECO2124,2125","✓","econ"),
  ("ECO3323","การบริหารความเสี่ยงองค์กร",3,5,"เลือกกลุ่มอื่น–การเงิน","ECO2121","✓","fin"),
  ("ECO2128","ระเบียบวิธีวิจัยทางเศรษฐศาสตร์",3,5,"แกน • ต่อยอด research paper","ECO2123-2126","✓","econ"),
  ("MTH4404","Introduction to Real Analysis 1",3,9,"เก็บเพิ่ม • สัญญาณ #1","MTH2104(+MTH2001)","✓","quant"),
  ("ACC4209","Financial Reporting & Analysis",3,6,"วิเคราะห์งบ (core IB)","ACC2202","✓","fin"),
  ("SDM2102","Foundations of Python Programming",3,4,"เก็บเพิ่ม • Python","-","✓","quant"),
  ("RAM1312","Contemporary Culture & Digital Disruption",3,2,"⬅ ย้ายมา v10 • ลดโหลด","-","✓","gen"),
 ]),
 ("ปีที่ 4 / เทอม 1",21,[
  ("ECO4124","Time-Series Econometrics",3,8,"เชี่ยวชาญ (เลือก)","ECO3123","✓","econ"),
  ("ECO4125","Microeconometrics",3,8,"เชี่ยวชาญ • เช็ก causal inf.","ECO3123","✓","econ"),
  ("ECO4326","Mathematical Finance",3,8,"เลือกกลุ่มอื่น–การเงิน","ECO3323","✓","fin"),
  ("MTH3611","Mathematical Stochastic Models",3,8,"เก็บเพิ่ม • Stochastic","MTH2106","✓ PR คณิต","quant"),
  ("FIN3211","Financial Management (Corp Finance)",3,6,"เก็บเพิ่ม • เนื้อ IB","FIN2101","⚠️ อาจต้อง FIN3208","fin"),
  ("RAM1204","Mathematics & Statistics for Daily Life",3,3,"⬅ ย้ายมา v9 • ศึกษาทั่วไป 1.2","-","✓","gen"),
  ("RAM1302","Politics and Law in Daily Life",3,2,"ศึกษาทั่วไป 1.3","-","✓","gen"),
 ]),
 ("ปีที่ 4 / เทอม 2",21,[
  ("ECO4123","Game Theory (ทฤษฎีเกม)",3,7,"เชี่ยวชาญ (เลือก)","ECO2123,2125","✓","econ"),
  ("ECO4327","Financial Risk Mgmt & Derivatives",3,7,"เลือกกลุ่มอื่น–การเงิน","ECO3320,3323","✓","fin"),
  ("COS2102","Object-Oriented Programming (C++)",3,6,"เก็บเพิ่ม • C++ (quant dev)","COS1103","✓","quant"),
  ("FIN3209","Securities Analysis",3,6,"เก็บเพิ่ม • valuation","FIN2101","✓","fin"),
  ("FIN3205","Principles & Policies of Investment",3,5,"เก็บเพิ่ม • การลงทุน","FIN2101","✓","fin"),
  ("RAM1211","Principles of Entrepreneurship",3,2,"ศึกษาทั่วไป 1.2","-","✓","gen"),
  ("STA4101","Theory of Statistics 1",3,9,"★ Math Stats (MFE core)","STA3102","⬅ ย้ายจากปี3/ท2 ลดโหลด","quant"),
 ]),
]

wb=Workbook(); ws=wb.active; ws.title="แผน v10 (ความยาก)"; ws.sheet_view.showGridLines=False
NC=8
ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=NC)
c=ws.cell(1,1,"แผนการเรียน ป.ตรี เศรษฐศาสตร์ ม.รามคำแหง — Quant+Finance/IB (v10: 168 นก. + ความยาก/10)")
c.font=FONT_TITLE; c.fill=TITLE_FILL; c.alignment=CENTER; ws.row_dimensions[1].height=24
ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=NC)
c=ws.cell(2,1,"เพิ่มจาก v9: คอลัมน์ 'ยาก/10' (heatmap) + 'สถานะบุพวิชา' • ขยับ STA4101 (9/10) ปี3/ท2 → ปี4/ท2 แยกจาก Real Analysis (เทอมหนักสุด 6.6→5.6) • บุพวิชาเรียงครบ เหลือยืนยัน ECO2126↔STA2003")
c.font=FONT_NOTE; c.alignment=LEFT; ws.row_dimensions[2].height=42
ws.merge_cells(start_row=3,start_column=1,end_row=3,end_column=NC)
c=ws.cell(3,1,"สีหมวด: ฟ้า=เศรษฐ เขียว=คณิต/สถิติ/โปรแกรม ส้ม=การเงิน/บัญชี เทา=ศึกษาทั่วไป | คอลัมน์ยาก/10: เขียว≤4 เหลือง5-6 ส้ม7-8 แดง9-10 | ⬅=ย้ายเทอม")
c.font=FONT_NOTE; c.alignment=LEFT
r=5
for j,h in enumerate(HEADERS,1):
    cell=ws.cell(r,j,h); cell.fill=HEADER_FILL; cell.font=FONT_HDR
    cell.alignment=CENTER if j in (1,2,4,5) else LEFT; cell.border=BORDER
ws.freeze_panes=ws.cell(r+1,1); r+=1
for term,total,rows in PLAN:
    avg=round(sum(x[3] for x in rows)/len(rows),1)
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC-1)
    tc=ws.cell(r,1,"%s   (ยากเฉลี่ย %s/10)"%(term,avg)); tc.fill=TERM_FILL; tc.font=FONT_B; tc.alignment=LEFT; tc.border=BORDER
    for j in range(2,NC): ws.cell(r,j).fill=TERM_FILL; ws.cell(r,j).border=BORDER
    cc=ws.cell(r,NC,"%d นก."%total); cc.fill=TERM_FILL; cc.font=FONT_B; cc.alignment=CENTER; cc.border=BORDER
    r+=1
    for code,name,cr,diff,reason,pre,status,cat in rows:
        vals=["",code,name,cr,diff,reason,pre,status]
        for j,v in enumerate(vals,1):
            cell=ws.cell(r,j,v); cell.fill=FILL[cat]; cell.font=FONT; cell.border=BORDER
            cell.alignment=CENTER if j in (1,2,4,5) else LEFT
        ws.cell(r,5).fill=dfill(diff); ws.cell(r,5).font=FONT_B   # heatmap on /10
        r+=1
r+=1
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC)
c=ws.cell(r,1,"สรุป: บุพวิชาเรียงครบทุกวิชา ✓ • เหลือ 'ต้องยืนยัน' = ECO2126↔STA2003 (สำคัญสุด), FIN2101/FIN3211/ACC2202 (ข้ามคณะ) • 5 วิชายากสุด: MTH4404, STA4101 (9) · MTH2104, STA3102, ECO3121/3122, ECO4124/4125/4326, MTH3611 (8)")
c.font=FONT_B; c.fill=TERM_FILL; c.alignment=LEFT; c.border=BORDER; ws.row_dimensions[r].height=30
for j,w in enumerate([4,9,40,5,8,36,22,30],1):
    ws.column_dimensions[get_column_letter(j)].width=w
wb.save(OUT); print("SAVED",OUT,"bytes=",os.path.getsize(OUT))
