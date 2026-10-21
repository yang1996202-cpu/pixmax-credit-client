#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PIXMAX 积分消耗分析器
  输入：credit-transactions CSV（自动识别「企业版 / 团队版」两种导出格式）
  输出：--view client    给客户/老板看的 Excel（9 表）
        --view business  给 PIXMAX 内部看的客户经营 Excel（5 表）

用法：
  python3 analyze.py <csv路径> --view client   [--out <输出目录>]
  python3 analyze.py <csv路径> --view business [--out <输出目录>]
"""
import argparse, csv, os, sys
from collections import defaultdict, Counter
from datetime import date, timedelta

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import RichTextProperties, Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule

# ============ 字段自动探测 ============
def detect(cols):
    def find(*keys):
        for k in keys:
            if k in cols: return k
        for k in keys:
            for c in cols:
                if k in c: return c
        return None
    f = {
        "time":   find("时间"),
        "kind":   find("类型"),
        "points": find("变动积分", "积分变化"),
        "person": find("账号姓名", "子账号ID", "子账号"),
        "model":  find("AI模型"),
        "amount": find("消耗积分对应金额款项"),
        "company":find("企业名称"),
        "balance":find("总余额", "积分余额"),
        "ptype":  find("积分类型"),
        "project":find("所属项目"),
    }
    missing = [k for k in ("time","kind","points","person") if not f[k]]
    if missing:
        raise SystemExit(f"[错误] 关键字段缺失：{missing}；实际列：{cols}")
    return f

# ============ 样式 ============
def xl(c): return "FF"+c.removeprefix("#").upper()
PRIMARY=xl("#2F6FBF"); HEAD=xl("#FFFFFF"); INPUT=xl("#DCE7F5"); SUM=xl("#1F4E79")
BD=xl("#BFBFBF"); ZB=xl("#F3F7FC"); DANGER=xl("#C00000"); WARNBG=xl("#FCE4E4")
GREEN=xl("#548235"); GREY=xl("#808080"); ORANGE=xl("#ED7D31"); TEAL=xl("#2E9CA6")
thin=Side(style="thin",color=BD)
F_TITLE=Font(bold=True,size=15,color=xl("#1F3864")); F_SUB=Font(bold=True,size=11,color=xl("#1F3864"))
F_NOTE=Font(size=9,color=GREY); F_HEAD=Font(bold=True,color=HEAD); F_SUMF=Font(bold=True,color=HEAD)
F_B=Font(bold=True)
FH=PatternFill("solid",fgColor=PRIMARY); FS=PatternFill("solid",fgColor=SUM)
FIN=PatternFill("solid",fgColor=INPUT); FZ=PatternFill("solid",fgColor=ZB); FW=PatternFill("solid",fgColor=WARNBG)
AC=Alignment("center","center"); AL=Alignment("left","center"); AR=Alignment("right","center")

def hrow(ws,row,n):
    for j in range(1,n+1):
        c=ws.cell(row=row,column=j); c.font=F_HEAD; c.fill=FH; c.alignment=AC; c.border=Border(bottom=thin)
def ttl(ws,t,span):
    ws["A1"]=t; ws["A1"].font=F_TITLE; ws["A1"].alignment=AC; ws.merge_cells(f"A1:{span}1")
def note(ws,row,text,col=1):
    ws.cell(row=row,column=col,value=text).font=F_NOTE
def labels(ch,sz=800,bold=False):
    dl=DataLabelList(); dl.showVal=True; dl.showCatName=False; dl.showSerName=False
    dl.showLegendKey=False; dl.showPercent=False; dl.showBubbleSize=False
    dl.txPr=RichText(bodyPr=RichTextProperties(),p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=sz,b=bold)))])
    ch.dataLabels=dl; ch.legend=None; ch.y_axis.majorGridlines=ChartLines()

# ============ 模型分类 ============
PIC={"GPT_IMAGE_2_5_FLARE","GPT_IMAGE_2_5_SUNBURST","GPT_IMAGE_2","DOUBAO_SEEDREAM_5_PRO","BANANA_2","GEMINI_31_PRO","QWEN_IMAGE_3_PRO"}
UPS={"IMAGE_UPSCALE","VIDEO_UPSCALE"}
def cat(m):
    if not m: return "-"
    if m in UPS: return "超分"
    if m in PIC: return "图像生成"
    return "视频生成"

# ============ 载入 + 聚合 ============
def load(path):
    rows=list(csv.DictReader(open(path,encoding="utf-8-sig")))
    if not rows: raise SystemExit("[错误] CSV 无数据")
    F=detect(list(rows[0].keys()))
    cons=[r for r in rows if (r[F["kind"]] or "").strip()=="消费"]
    rech=[r for r in rows if (r[F["kind"]] or "").strip()!="消费"]
    def pv(r):
        try: return int(float(r[F["points"]] or 0))
        except: return 0
    def amt(r):
        if not F["amount"]: return 0.0
        try: return float(r[F["amount"]] or 0)
        except: return 0.0
    for r in cons+rech: r["_d"]=(r[F["time"]] or "")[:10]
    D={"F":F,"cons":cons,"rech":rech,"pv":pv,"amt":amt,"rows":rows}
    if not cons: raise SystemExit("[错误] 无「消费」记录")
    D["company"]=(cons[0][F["company"]] if F["company"] else "（未提供企业名）")
    # 人（子账号ID 为空 = 主账号直用，统一记「主账号」，不再显示「（未知）」）
    MAIN_TAG="主账号"
    person=defaultdict(lambda:{"积分":0,"次数":0,"金额":0.0,"模型":Counter(),"cat":Counter()})
    for r in cons:
        p=person[(r[F["person"]] or "").strip() or MAIN_TAG]
        p["积分"]+=pv(r); p["次数"]+=1; p["金额"]+=amt(r)
        if F["model"]: p["模型"][r[F["model"]]]+=1
        p["cat"][cat(r[F["model"]]) if F["model"] else "视频生成"]+=1
    D["person"]=person; D["people"]=sorted(person,key=lambda k:-person[k]["积分"])
    # 人员键归一：回填到原始行，保证后续按人过滤时口径一致
    for r in cons+rech: r["_p"]=(r[F["person"]] or "").strip() or MAIN_TAG
    # 无子账号场景标记（企业版导出但未开子账号，全记主账号名下）
    D["no_sub"]=all((r[F["person"]] or "").strip()=="" for r in cons)
    D["TOT_JI"]=sum(person[k]["积分"] for k in D["people"])
    D["TOT_AMT"]=sum(person[k]["金额"] for k in D["people"])
    D["TOT_CNT"]=len(cons)
    D["has_amt"]=bool(F["amount"]); D["has_model"]=bool(F["model"])
    # 模型
    model=defaultdict(lambda:{"积分":0,"次数":0,"金额":0.0,"用户":set(),"档位":Counter()})
    for r in cons:
        m=model[(r[F["model"]] if F["model"] else "（无模型信息）") or "(空)"]
        m["积分"]+=pv(r); m["次数"]+=1; m["金额"]+=amt(r)
        m["用户"].add(r[F["person"]]); m["档位"][pv(r)]+=1
    D["model"]=model; D["models"]=sorted(model,key=lambda k:-model[k]["积分"])
    # 天/周
    day=defaultdict(lambda:{"积分":0,"次数":0,"金额":0.0,"用户":set()})
    for r in cons:
        d=day[r["_d"]]; d["积分"]+=pv(r); d["次数"]+=1; d["金额"]+=amt(r); d["用户"].add(r[F["person"]])
    D["day"]=day; D["days"]=sorted(day)
    week=defaultdict(lambda:{"积分":0,"次数":0,"金额":0.0,"日":set()})
    for r in cons:
        d=date(*map(int,r["_d"].split("-"))); wk=(d-timedelta(days=d.weekday())).isoformat()
        w=week[wk]; w["积分"]+=pv(r); w["次数"]+=1; w["金额"]+=amt(r); w["日"].add(r["_d"])
    D["week"]=week; D["weeks"]=sorted(week)
    # 余额/充值
    if F["balance"]:
        try: D["BALANCE"]=int(float(sorted(D["rows"],key=lambda r:r[F["time"]])[-1][F["balance"]] or 0))
        except: D["BALANCE"]=0
    else: D["BALANCE"]=0
    if F["ptype"]:
        D["PAID"]=sum(pv(r) for r in rech if (r[F["ptype"]] or "")=="活动充值积分")
        D["POOL"]=sum(pv(r) for r in rech)
        ck=defaultdict(int)
        for r in cons: ck[r[F["ptype"]]]+=pv(r)
        D["PAID_USED"]=ck.get("活动充值积分",0)
    else:
        D["PAID"]=0; D["POOL"]=sum(pv(r) for r in rech) or D["TOT_JI"]; D["PAID_USED"]=0
    # 节奏
    d0=date(*map(int,D["days"][0].split("-"))); d1=date(*map(int,D["days"][-1].split("-")))
    D["d0"]=d0; D["d1"]=d1; D["span"]=(d1-d0).days+1
    D["r7"]=sum(day.get((d1-timedelta(days=i)).isoformat(),{"积分":0})["积分"] for i in range(7))
    D["r7avg"]=D["r7"]/7
    D["avg_all"]=D["TOT_JI"]/D["span"]
    D["left"]=D["BALANCE"]/D["r7avg"] if D["r7avg"] else 0
    D["eta"]=(d1+timedelta(days=round(D["left"]))).isoformat()
    D["peak"]=max(D["days"],key=lambda d:day[d]["积分"])
    return D

WD=["周一","周二","周三","周四","周五","周六","周日"]
def wd(s):
    y,m,d=map(int,s.split("-")); return WD[date(y,m,d).weekday()]

# ============ 视角一：客户 / 老板 ============
def build_client(D,out):
    F=D["F"]; people=D["people"]; person=D["person"]
    TOT_JI=D["TOT_JI"]; TOT_AMT=D["TOT_AMT"]; TOT_CNT=D["TOT_CNT"]
    days=D["days"]; day=D["day"]; weeks=D["weeks"]; week=D["week"]
    models=D["models"]; model=D["model"]; BALANCE=D["BALANCE"]
    PAID=D["PAID"]; POOL=D["POOL"]; PAID_USED=D["PAID_USED"]; r7avg=D["r7avg"]; left=D["left"]
    has_amt=D["has_amt"]; has_model=D["has_model"]
    ENT=D["company"]
    # 增量：高配/省钱
    hi=[r for r in D["cons"] if D["pv"](r)>=400]
    HI_N=len(hi); HI_J=sum(D["pv"](r) for r in hi)
    hi_ji=defaultdict(int); hi_cnt=Counter()
    for r in hi: hi_ji[r["_p"]]+=D["pv"](r); hi_cnt[r["_p"]]+=1
    p25=[r for r in D["cons"] if has_model and r[F["model"]]=="PIXDANCE_2_5"]
    J25=sum(D["pv"](r) for r in p25)
    # 降级省钱口径：FAST 实测单次均价优先（无 FAST 数据时退回 90 分/次经验值）
    fast_rows=[r for r in D["cons"] if has_model and r[F["model"]]=="PIXDANCE_2_FAST"]
    FAST_AVG=(sum(D["pv"](r) for r in fast_rows)/len(fast_rows)) if fast_rows else 90
    FAST_AVG=FAST_AVG if fast_rows else 90
    SAVE=round(J25-len(p25)*FAST_AVG); SAVE_AMT=SAVE*0.0266
    # 周环比加速：取最近完整两周对比（前期→近期），而非写死第3周/第2周
    if len(weeks)>2 and week[weeks[-2]]["积分"]:
        w_prev,w_last=weeks[-2],weeks[-1]
        spike=week[w_last]["积分"]/week[w_prev]["积分"]
        spike_txt=f"第{len(weeks)-1}周→第{len(weeks)}周为 {spike:.1f} 倍"
    elif len(weeks)>1 and week[weeks[0]]["积分"]:
        w_prev,w_last=weeks[0],weeks[1]
        spike=week[w_last]["积分"]/week[w_prev]["积分"]
        spike_txt=f"第1周→第2周为 {spike:.1f} 倍"
    else:
        spike=0; spike_txt="-"
    last2=(week[weeks[-1]]["积分"]+week[weeks[-2]]["积分"]) if len(weeks)>1 else TOT_JI

    wb=Workbook()
    # ---- 老板重点 ----
    ws=wb.active; ws.title="老板重点"
    ttl(ws,f"老板重点速览 · {ENT}（{days[0]} ~ {days[-1]}）","E")
    r=2
    d0=date(*map(int,days[0].split("-"))); d1=date(*map(int,days[-1].split("-")))
    span=(d1-d0).days+1
    used_days=len(day); gap=span-used_days
    for lab,val in [("消耗区间",f"{days[0]} → {days[-1]}"),
                    ("覆盖天数",f"{span} 天（有消耗 {used_days} 天 / 断档 {gap} 天）"),
                    ("数据截止",f"{days[-1]}（导出请核对是否含最新日期）")]:
        a=ws.cell(row=r,column=1,value=lab); a.font=F_B; a.fill=FIN; a.border=Border(bottom=thin); a.alignment=AC
        b=ws.cell(row=r,column=2,value=val); b.font=Font(bold=True,size=11,color=xl("#1F3864")); b.alignment=AL; ws.merge_cells(f"B{r}:E{r}")
        for j2 in range(1,6): ws.cell(row=r,column=j2).border=Border(bottom=thin)
        r+=1
    r+=1
    c=ws.cell(row=r,column=1,value=f"结论：剩余 {BALANCE:,} 积分，按近7天速度约够 {left:.1f} 天；累计消耗 {TOT_JI:,} 积分"+(f"（¥{TOT_AMT:,.2f}）" if has_amt else "")+f"，共 {TOT_CNT:,} 次调用。")
    c.font=Font(bold=True,size=12,color=DANGER); c.fill=FW; c.alignment=AL; ws.merge_cells(f"A{r}:E{r}")
    r+=2
    ws.cell(row=r,column=1,value="① 你最该关注的三件事").font=F_SUB; r+=1
    for j,h in enumerate(["关注点","数字","说明"],2): ws.cell(row=r,column=j,value=h)
    hrow(ws,r,3); ws.cell(row=r,column=4,value="说明"); ws.merge_cells(f"D{r}:E{r}")
    # 三件事·谁花得最多：无子账号场景显示主账号而非「（未知）」
    top_name="主账号（未开子账号）" if D.get("no_sub") and people[0]=="主账号" else people[0]
    three=[("还能烧多久",f"{left:.1f} 天",f"剩余 {BALANCE:,} 积分 ÷ 近7天日均 {round(r7avg):,} 积分/天"),
           ("谁花得最多",f"{top_name} · "+(f"¥{person[people[0]]['金额']:,.0f}" if has_amt else f"{person[people[0]]['积分']:,}分"),f"占总量 {person[people[0]]['积分']/TOT_JI*100:.1f}%，高配调用 {hi_cnt.get(people[0],0)} 笔（≥400分）"),
           ("理论上能省多少",(f"¥{SAVE_AMT:,.0f}（{SAVE/TOT_JI*100:.0f}%）" if has_amt else f"{SAVE:,} 积分（{SAVE/TOT_JI*100:.0f}%）"),(f"PIXDANCE_2_5 的 {len(p25)} 次调用降级到 FAST 可省 {SAVE:,} 积分" if has_model else "需含 AI模型 字段方可用"))]
    for i,(a,b,cc) in enumerate(three):
        rr=r+1+i
        ws.cell(row=rr,column=2,value=a).font=F_B
        y=ws.cell(row=rr,column=3,value=b); y.font=Font(bold=True,color=DANGER); y.alignment=AR
        z=ws.cell(row=rr,column=4,value=cc); z.alignment=AL; ws.merge_cells(f"D{rr}:E{rr}")
        for j in range(2,6): ws.cell(row=rr,column=j).border=Border(bottom=thin)
    r=r+4
    ws.cell(row=r,column=1,value="② 异常值清单").font=F_SUB; r+=1
    for j,h in enumerate(["#","异常现象","涉及账号","数据","疑似原因"],1): ws.cell(row=r,column=j,value=h)
    hrow(ws,r,5)
    exts=sorted([r_ for r_ in D["cons"] if D["pv"](r_)>=800],key=lambda r_:-D["pv"](r_))
    anom=[("1","单次消耗破纪录",((exts[0]["_p"] if exts else "-")+" "+exts[0]["_d"] if exts else "-"),
           (f"{len(exts)} 笔 ≥800 分，最高 {D['pv'](exts[0]):,} 分" if exts else "无"),("单次成本高（数据事实）" if exts else "无异常")),
          ("2","高价调用集中度",top_name,
           f"{hi_cnt.get(people[0],0)} 笔 ≥400 分（{hi_ji.get(people[0],0):,} 分），占全部高配 {hi_ji.get(people[0],0)/HI_J*100:.0f}%" if HI_J else "无 ≥400 分调用","仅陈述分布，不做归因"),
          ("3",("贵模型占比" if has_model else "N/A"),("全体" if has_model else "N/A"),
           (f"PIXDANCE_2_5 仅 {len(p25)} 次却占 {J25/TOT_JI*100:.0f}% 积分" if has_model else "无模型字段"),"高价档调用，成本影响见上行"),
          ("4",("消耗突然加速" if spike>1.3 else "消耗节奏变化"),"全体",(spike_txt if spike else "-"),("周环比明显上升（数据事实）" if spike>1.3 else ("节奏平稳/回落（数据事实）" if spike else "-"))),
          ("5","高配调用总量",f"{HI_N} 笔",f"占全部调用 {HI_N/TOT_CNT*100:.1f}%，却消耗 {HI_J/TOT_JI*100:.1f}% 积分" if HI_N else "无高配调用","少数调用占大头积分（数据事实）")]
    for i,row_ in enumerate(anom):
        rr=r+1+i
        for j,v in enumerate(row_,1):
            cell=ws.cell(row=rr,column=j,value=v); cell.alignment=AC if j==1 else AL
            cell.border=Border(bottom=thin)
            if j==5: cell.font=Font(color=DANGER)
    r=r+len(anom)+2
    ws.cell(row=r,column=1,value="③ 原因拆解").font=F_SUB; r+=1
    c=ws.cell(row=r,column=2,value=f"用量在涨（最近两周占 {last2/TOT_JI*100:.0f}%） × 用贵模型（{J25/TOT_JI*100:.0f}% 积分来自 PIXDANCE_2_5） × 用高配参数（{HI_J/TOT_JI*100:.0f}% 积分来自 {HI_N} 笔高配）" if has_model else f"用量在涨（最近两周占 {last2/TOT_JI*100:.0f}%）× 高配参数（{HI_J/TOT_JI*100:.0f}% 积分来自 {HI_N} 笔高配调用）")
    c.font=F_B; c.alignment=AL; ws.merge_cells(f"B{r}:E{r}")
    r+=2
    if has_model:
        s7d=(d1-timedelta(days=6)).isoformat()
        m7=defaultdict(lambda:{"积分":0,"次数":0}); mpre=defaultdict(lambda:{"积分":0,"次数":0})
        for r_ in D["cons"]:
            if not r_.get(F["model"]) or not r_[F["model"]]: continue
            tgt=m7 if r_["_d"]>=s7d else mpre
            tgt[r_[F["model"]]]["积分"]+=D["pv"](r_); tgt[r_[F["model"]]]["次数"]+=1
        J7=sum(v["积分"] for v in m7.values()); Jpre=sum(v["积分"] for v in mpre.values())
        C7=sum(v["次数"] for v in m7.values()); Cpre=sum(v["次数"] for v in mpre.values())
        if J7 or Jpre:
            u7=J7/C7 if C7 else 0; upre=Jpre/Cpre if Cpre else 0
            trend="档位平稳" if (not upre or abs(u7-upre)/upre<=0.10) else ("单价上探（转向更贵配置）" if u7>upre else "单价下探（转向更便宜配置）")
            mset=[m for m in sorted(set(m7)|set(mpre),key=lambda m:-m7[m]["积分"]) if m7[m]["积分"]>0 or (Jpre and mpre[m]["积分"]/Jpre*100>=1.0)]
            top7=mset[0] if mset else "-"
            ws.cell(row=r,column=1,value=f"④ 近7天模型动态（{s7d[5:]} ~ {days[-1][5:]} vs 前期）").font=F_SUB; r+=1
            c=ws.cell(row=r,column=1,value=f"近7天共耗 {J7:,} 积分 / {C7:,} 次；单次均价 {u7:.0f} 分 vs 前期 {upre:.0f} 分 → {trend}。主力：{top7} 占近7天 {(m7[top7]['积分']/J7*100 if J7 and top7!='-' else 0):.0f}%。")
            c.font=F_B; c.alignment=AL; ws.merge_cells(f"A{r}:E{r}"); r+=2
            for j,h in enumerate(["AI模型","近7天积分","近7天占比","前期占比","占比变化"],1): ws.cell(row=r,column=j,value=h)
            hrow(ws,r,5)
            for i,m in enumerate(mset):
                rr2=r+1+i; s7v=m7[m]["积分"]; spv=mpre[m]["积分"]
                p7=s7v/J7*100 if J7 else 0; pp=spv/Jpre*100 if Jpre else 0
                ws.cell(row=rr2,column=1,value=m)
                ws.cell(row=rr2,column=2,value=s7v).number_format="#,##0"
                for col,v in [(3,p7),(4,pp),(5,p7-pp)]:
                    cc=ws.cell(row=rr2,column=col,value=f"{v:+.1f}%" if col==5 else f"{v:.1f}%"); cc.alignment=AR
                    cc.font=Font(color=DANGER) if p7>pp else (Font(color=GREEN) if p7<pp else F_NOTE)
                for j in range(1,6): ws.cell(row=rr2,column=j).border=Border(bottom=thin)
            r=r+1+len(mset)+1
    ws.cell(row=r,column=1,value="⑤ 建议动作").font=F_SUB; r+=1
    acts=[f"1. 提醒客户充值（按当前速度 {left:.1f} 天见底）。"]
    if HI_N: acts.append(f"2. 排查 {top_name} 的高配调用（≥400 分共 {HI_N} 笔）是否都必要。")
    if exts: acts.append(f"3. 核实单次 ≥800 分的调用（共 {len(exts)} 笔，最高 {D['pv'](exts[0]):,} 分）是什么任务。")
    if has_model and J25 and len(acts)<4:
        n=str(len(acts)+1)
        acts.append(f"{n}. 高价档 PIXDANCE_2_5 占 {J25/TOT_JI*100:.0f}% 积分，可评估非必要场景降级。")
    if len(acts)<2: acts.append("2. 用量结构正常，按常规节奏跟进即可。")
    for i,a in enumerate(acts):
        c=ws.cell(row=r+i,column=2,value=a); c.alignment=AL; ws.merge_cells(f"B{r+i}:E{r+i}")
    for j,w_ in enumerate([4,22,20,34,40],1): ws.column_dimensions[get_column_letter(j)].width=w_

    # ---- 总览 ----
    ws=wb.create_sheet("总览")
    ttl(ws,f"积分消耗分析总览 · {ENT}","J")
    kpi=[("统计周期",f"{days[0]} ~ {days[-1]}（{len(days)}天有消耗）",None),
         ("总消耗积分",TOT_JI,"#,##0"),("总花费（元）",round(TOT_AMT,2),"¥#,##0.00") if has_amt else ("总花费（元）","无金额字段",None),
         ("总调用次数（生成任务数）",TOT_CNT,"#,##0"),(("活跃子账号" if F["person"]!="账号姓名" else "活跃操作用户"),len(people),"0"),
         ("日均消耗积分",round(TOT_JI/len(days)),"#,##0"),("峰值消耗日",f"{D['peak']}（{wd(D['peak'])}）",None),("剩余积分",BALANCE,"#,##0")]
    r=3
    for lab,val,fmt in kpi:
        a=ws.cell(row=r,column=1,value=lab); a.font=F_B; a.fill=FIN; a.border=Border(bottom=thin)
        b=ws.cell(row=r,column=2,value=val); b.font=F_B; b.alignment=AR; b.border=Border(bottom=thin)
        if fmt: b.number_format=fmt
        r+=1
    a=ws.cell(row=r,column=1,value="⚠ 预计可支撑"); a.font=Font(bold=True,size=12,color=DANGER); a.fill=FW
    b=ws.cell(row=r,column=2,value=f"约 {left:.1f} 天"); b.font=Font(bold=True,size=13,color=DANGER); b.fill=FW; b.alignment=AR
    r+=1
    note(ws,r,"口径：调用次数 = 生成任务数（一次任务计 1 次）；单次积分 = 消耗积分 ÷ 调用次数（积分/次，非金额）")
    wr=r+2
    ws.cell(row=wr,column=1,value="按周消耗趋势").font=F_SUB
    for j,h in enumerate(["周次","日期区间","消耗积分","调用次数(次)","花费(元)","占总比"],1): ws.cell(row=wr+1,column=j,value=h)
    hrow(ws,wr+1,6)
    for i,wk in enumerate(weeks):
        rr=wr+2+i; w=week[wk]; end=(date(*map(int,wk.split("-")))+timedelta(days=6)).isoformat()
        ws.cell(row=rr,column=1,value=f"第{i+1}周").alignment=AC
        ws.cell(row=rr,column=2,value=f"{wk[5:]} ~ {end[5:]}"+("" if len(w["日"])==7 else f"（{len(w['日'])}天）")).alignment=AC
        ws.cell(row=rr,column=3,value=w["积分"]).number_format="#,##0"
        ws.cell(row=rr,column=4,value=w["次数"]).number_format="#,##0"
        ws.cell(row=rr,column=5,value=round(w["金额"],2) if has_amt else "-").number_format="¥#,##0.00"
        ws.cell(row=rr,column=6,value=f"{w['积分']/TOT_JI*100:.1f}%").alignment=AC
        for j in range(1,7): ws.cell(row=rr,column=j).border=Border(bottom=thin)
    wsum=wr+2+len(weeks)
    ws.cell(row=wsum,column=1,value="合计")
    ws.cell(row=wsum,column=3,value=TOT_JI).number_format="#,##0"
    ws.cell(row=wsum,column=4,value=TOT_CNT).number_format="#,##0"
    ws.cell(row=wsum,column=6,value="100.0%").alignment=AC
    for j in range(1,7):
        c=ws.cell(row=wsum,column=j); c.font=F_SUMF; c.fill=FS; c.alignment=AC
    ws.conditional_formatting.add(f"C{wr+2}:C{wsum-1}",DataBarRule(start_type="num",start_value=0,end_type="max",color=PRIMARY,showValue=True))
    for j,w_ in enumerate([24,17,12,14,12,10],1): ws.column_dimensions[get_column_letter(j)].width=w_

    # ---- 按天汇总 ----
    ws2=wb.create_sheet("按天汇总"); ttl(ws2,"按天消耗趋势","G")
    for j,h in enumerate(["日期","星期","所属周","消耗积分","调用次数(次)","花费(元)","占总比"],1): ws2.cell(row=2,column=j,value=h)
    hrow(ws2,2,7); f0=3
    for i,d in enumerate(days):
        rr=f0+i; a=day[d]; dd=date(*map(int,d.split("-"))); wk=(dd-timedelta(days=dd.weekday())).isoformat()
        ws2.cell(row=rr,column=1,value=d).alignment=AC; ws2.cell(row=rr,column=2,value=wd(d)).alignment=AC
        ws2.cell(row=rr,column=3,value=wk[5:]).alignment=AC
        ws2.cell(row=rr,column=4,value=a["积分"]).number_format="#,##0"
        ws2.cell(row=rr,column=5,value=a["次数"]).number_format="#,##0"
        ws2.cell(row=rr,column=6,value=round(a["金额"],2) if has_amt else "-").number_format="¥#,##0.00"
        ws2.cell(row=rr,column=7,value=f"{a['积分']/TOT_JI*100:.1f}%").alignment=AC
        for j in range(1,8): ws2.cell(row=rr,column=j).border=Border(bottom=thin)
    last=f0+len(days)-1; sr=last+1
    ws2.cell(row=sr,column=1,value="合计").alignment=AC
    ws2.cell(row=sr,column=4,value=TOT_JI).number_format="#,##0"
    ws2.cell(row=sr,column=5,value=TOT_CNT).number_format="#,##0"
    ws2.cell(row=sr,column=7,value="100.0%").alignment=AC
    for j in range(1,8):
        c=ws2.cell(row=sr,column=j); c.font=F_SUMF; c.fill=FS; c.alignment=AC
    ws2.conditional_formatting.add(f"D{f0}:D{last}",DataBarRule(start_type="num",start_value=0,end_type="max",color=PRIMARY,showValue=True))
    for j,w_ in enumerate([12,8,9,12,14,12,10],1): ws2.column_dimensions[get_column_letter(j)].width=w_
    ws2.freeze_panes="A3"

    # ---- 按周汇总 ----
    ws3=wb.create_sheet("按周汇总"); ttl(ws3,"按周消耗趋势","F")
    for j,h in enumerate(["周次","日期区间","消耗积分","调用次数(次)","花费(元)","占总比"],1): ws3.cell(row=2,column=j,value=h)
    hrow(ws3,2,6); f0=3
    for i,wk in enumerate(weeks):
        rr=f0+i; w=week[wk]; end=(date(*map(int,wk.split("-")))+timedelta(days=6)).isoformat()
        ws3.cell(row=rr,column=1,value=f"第{i+1}周").alignment=AC
        ws3.cell(row=rr,column=2,value=f"{wk} ~ {end}"+("" if len(w["日"])==7 else f"（{len(w['日'])}天）")).alignment=AC
        ws3.cell(row=rr,column=3,value=w["积分"]).number_format="#,##0"
        ws3.cell(row=rr,column=4,value=w["次数"]).number_format="#,##0"
        ws3.cell(row=rr,column=5,value=round(w["金额"],2) if has_amt else "-").number_format="¥#,##0.00"
        ws3.cell(row=rr,column=6,value=f"{w['积分']/TOT_JI*100:.1f}%").alignment=AC
        for j in range(1,7): ws3.cell(row=rr,column=j).border=Border(bottom=thin)
    last=f0+len(weeks)-1; sr=last+1
    ws3.cell(row=sr,column=1,value="合计").alignment=AC
    ws3.cell(row=sr,column=3,value=TOT_JI).number_format="#,##0"
    ws3.cell(row=sr,column=4,value=TOT_CNT).number_format="#,##0"
    ws3.cell(row=sr,column=6,value="100.0%").alignment=AC
    for j in range(1,7):
        c=ws3.cell(row=sr,column=j); c.font=F_SUMF; c.fill=FS; c.alignment=AC
    for j,w_ in enumerate([9,26,12,14,12,10],1): ws3.column_dimensions[get_column_letter(j)].width=w_

    # ---- 按人汇总 ----
    pname="操作人" if F["person"]=="账号姓名" else "子账号ID"
    if D.get("no_sub"): pname="账号"  # 无子账号场景，列头不再叫「子账号ID」
    # 无子账号时行名「主账号」替换为「主账号（未开子账号）」
    disp=lambda s: ("主账号（未开子账号）" if s=="主账号" and D.get("no_sub") else s)
    ws4=wb.create_sheet("按人汇总"); ttl(ws4,f"按人（{pname}）消耗汇总 —— 人效成本侧","K")
    heads=["排名",pname,"消耗积分","调用次数(次)","花费(元)","单次积分(积分/次)","主力模型(调用最多)","视频类(次)","图片类(次)","超分(次)","占总比"]
    for j,h in enumerate(heads,1): ws4.cell(row=2,column=j,value=h)
    hrow(ws4,2,len(heads)); f0=3
    for i,s in enumerate(people):
        rr=f0+i; p=person[s]
        ws4.cell(row=rr,column=1,value=i+1).alignment=AC
        ws4.cell(row=rr,column=2,value=disp(s))
        ws4.cell(row=rr,column=3,value=p["积分"]).number_format="#,##0"
        ws4.cell(row=rr,column=4,value=p["次数"]).number_format="#,##0"
        ws4.cell(row=rr,column=5,value=round(p["金额"],2) if has_amt else "-").number_format="¥#,##0.00"
        ws4.cell(row=rr,column=6,value=round(p["积分"]/p["次数"],1)).number_format="0.0"
        ws4.cell(row=rr,column=7,value=(f"{p['模型'].most_common(1)[0][0]}（{p['模型'].most_common(1)[0][1]}次）" if p["模型"] else "-"))
        if has_model:
            ws4.cell(row=rr,column=8,value=p["cat"]["视频生成"]).number_format="#,##0"
            ws4.cell(row=rr,column=9,value=p["cat"]["图像生成"]).number_format="#,##0"
            ws4.cell(row=rr,column=10,value=p["cat"]["超分"]).number_format="#,##0"
        else:
            for col in (8,9,10): ws4.cell(row=rr,column=col,value="-").alignment=AC
        ws4.cell(row=rr,column=11,value=f"{p['积分']/TOT_JI*100:.1f}%").alignment=AC
        for j in range(1,12): ws4.cell(row=rr,column=j).border=Border(bottom=thin)
    last=f0+len(people)-1; sr=last+1
    ws4.cell(row=sr,column=2,value="合计")
    for col,val,fmt in [(3,TOT_JI,"#,##0"),(4,TOT_CNT,"#,##0"),(5,(round(TOT_AMT,2) if has_amt else "-"),"¥#,##0.00"),(6,round(TOT_JI/TOT_CNT,1),"0.0")]: ws4.cell(row=sr,column=col,value=val).number_format=fmt
    ws4.cell(row=sr,column=11,value="100.0%").alignment=AC
    if has_model:
        ws4.cell(row=sr,column=8,value=sum(person[s]["cat"]["视频生成"] for s in people)).number_format="#,##0"
        ws4.cell(row=sr,column=9,value=sum(person[s]["cat"]["图像生成"] for s in people)).number_format="#,##0"
        ws4.cell(row=sr,column=10,value=sum(person[s]["cat"]["超分"] for s in people)).number_format="#,##0"
    for j in range(1,12):
        c=ws4.cell(row=sr,column=j); c.font=F_SUMF; c.fill=FS; c.alignment=AC
    note(ws4,sr+2,"口径：单次积分 = 消耗积分 ÷ 调用次数；调用次数 = 生成任务数；主力模型 = 调用次数最多的模型")
    ws4.conditional_formatting.add(f"F{f0}:F{last}",ColorScaleRule(start_type="min",start_color="FFFFFFFF",end_type="max",end_color="FFE06666"))
    ws4.conditional_formatting.add(f"K{f0}:K{last}",DataBarRule(start_type="num",start_value=0,end_type="max",color=PRIMARY,showValue=True))
    for j,w_ in enumerate([6,13,12,13,12,15,22,13,13,12,10],1): ws4.column_dimensions[get_column_letter(j)].width=w_
    ws4.freeze_panes="A3"

    # ---- 人×模型 / 按模型（仅当有模型字段） ----
    if has_model:
        MAIN=[m for m in ["WAN3_0","PIXDANCE_2_5","PIXDANCE_2_FAST","PIXDANCE_2","PIXDANCE_2_MINI"] if m in model]
        def bucket(m):
            if m in MAIN: return "main"
            if m in PIC: return "pic"
            if m in UPS: return "ups"
            return "ov"
        ws5=wb.create_sheet("人×模型"); ttl(ws5,"谁在用什么模型（穿透视图）","I")
        mc=MAIN+["图片类","超分","其他视频","合计"] if MAIN else ["图片类","超分","其他视频","合计"]
        def build(ws5,top,kind):
            ws5.cell(row=top,column=1,value=kind).font=F_SUB
            for j,h in enumerate([pname]+mc,1): ws5.cell(row=top+1,column=j,value=h)
            hrow(ws5,top+1,len(mc)+1); d0=top+2; ncol=len(mc)+1
            for i,s in enumerate(people):
                rr=d0+i; ws5.cell(row=rr,column=1,value=disp(s)); mv=person[s]["模型"]
                out=[]
                for m in MAIN:
                    out.append(mv.get(m,0) if kind=="调用次数" else round(sum(D["amt"](x) for x in D["cons"] if x["_p"]==s and x[F["model"]]==m),2))
                pic=ups=ov=0
                for x in D["cons"]:
                    if x["_p"]!=s: continue
                    b=bucket(x[F["model"]])
                    val=1 if kind=="调用次数" else D["amt"](x)
                    if b=="pic": pic+=val
                    elif b=="ups": ups+=val
                    elif b=="ov": ov+=val
                out+=[pic if kind=="调用次数" else round(pic,2),ups if kind=="调用次数" else round(ups,2),ov if kind=="调用次数" else round(ov,2)]
                rowtot=sum(mv.values()) if kind=="调用次数" else round(sum(D["amt"](x) for x in D["cons"] if x["_p"]==s),2)
                for j,v in enumerate(out,start=2):
                    ws5.cell(row=rr,column=j,value=v if v else None)
                    ws5.cell(row=rr,column=j).number_format="#,##0" if kind=="调用次数" else "¥#,##0.00"
                ws5.cell(row=rr,column=2+len(mc)-1,value=rowtot)
                ws5.cell(row=rr,column=2+len(mc)-1).number_format="#,##0" if kind=="调用次数" else "¥#,##0.00"
                for j in range(1,ncol+1): ws5.cell(row=rr,column=j).border=Border(bottom=thin)
            dend=d0+len(people)-1; wsr=dend+1
            ws5.cell(row=wsr,column=1,value="合计")
            if kind=="调用次数": mt=[model[m]["次数"] for m in MAIN]; g=TOT_CNT
            else: mt=[round(model[m]["金额"],2) for m in MAIN]; g=round(TOT_AMT,2)
            pic_t=ups_t=ov_t=0
            for x in D["cons"]:
                b=bucket(x[F["model"]]); val=1 if kind=="调用次数" else D["amt"](x)
                if b=="pic": pic_t+=val
                elif b=="ups": ups_t+=val
                elif b=="ov": ov_t+=val
            if kind!="调用次数": pic_t,ups_t,ov_t=round(pic_t,2),round(ups_t,2),round(ov_t,2)
            for j,v in enumerate(mt+[pic_t,ups_t,ov_t,g],start=2):
                c=ws5.cell(row=wsr,column=j,value=v); c.font=F_SUMF; c.fill=FS; c.alignment=AC
                c.number_format="#,##0" if kind=="调用次数" else "¥#,##0.00"
            ws5.cell(row=wsr,column=1).font=F_SUMF; ws5.cell(row=wsr,column=1).fill=FS
            ws5.conditional_formatting.add(f"B{d0}:{get_column_letter(1+len(mc)-1)}{dend}",ColorScaleRule(start_type="min",start_color="FFFFFFFF",end_type="max",end_color="FFD64545"))
            return wsr
        e1=build(ws5,2,"调用次数"); e2=build(ws5,e1+2,"花费金额")
        note(ws5,e2+2,"分组：图片类 = 图像生成模型；超分 = IMAGE_UPSCALE / VIDEO_UPSCALE；其他视频 = 其余视频模型")
        for j,w_ in enumerate([13]+[12]*len(mc),1): ws5.column_dimensions[get_column_letter(j)].width=w_
        ws5.freeze_panes="B4"

        ws6=wb.create_sheet("按模型汇总"); ttl(ws6,"按模型消耗汇总（含类别 / 配置档位 / 超分）","J")
        heads=["排名","AI模型","类别","消耗积分","调用次数(次)","花费(元)","单次积分(积分/次)","配置档位数(种)","调用人数(人)","占总比"]
        for j,h in enumerate(heads,1): ws6.cell(row=2,column=j,value=h)
        hrow(ws6,2,len(heads)); f0=3
        for i,m in enumerate(models):
            rr=f0+i; d=model[m]; c0=cat(m)
            ws6.cell(row=rr,column=1,value=i+1).alignment=AC
            ws6.cell(row=rr,column=2,value=m)
            cc=ws6.cell(row=rr,column=3,value=c0); cc.alignment=AC
            if c0=="超分": cc.font=Font(bold=True,color=ORANGE)
            elif c0=="图像生成": cc.font=Font(color=TEAL)
            ws6.cell(row=rr,column=4,value=d["积分"]).number_format="#,##0"
            ws6.cell(row=rr,column=5,value=d["次数"]).number_format="#,##0"
            ws6.cell(row=rr,column=6,value=round(d["金额"],2) if has_amt else "-").number_format="¥#,##0.00"
            ws6.cell(row=rr,column=7,value=round(d["积分"]/d["次数"],1)).number_format="0.0"
            ws6.cell(row=rr,column=8,value=len(d["档位"])).alignment=AC
            ws6.cell(row=rr,column=9,value=len(d["用户"])).alignment=AC
            ws6.cell(row=rr,column=10,value=f"{d['积分']/TOT_JI*100:.1f}%").alignment=AC
            for j in range(1,11): ws6.cell(row=rr,column=j).border=Border(bottom=thin)
        last=f0+len(models)-1; sr=last+1
        ws6.cell(row=sr,column=2,value="合计")
        ws6.cell(row=sr,column=4,value=TOT_JI).number_format="#,##0"
        ws6.cell(row=sr,column=5,value=TOT_CNT).number_format="#,##0"
        ws6.cell(row=sr,column=10,value="100.0%").alignment=AC
        for j in range(1,11):
            c=ws6.cell(row=sr,column=j); c.font=F_SUMF; c.fill=FS; c.alignment=AC
        ur=sr+2
        ws6.cell(row=ur,column=2,value="超分增强使用情况").font=F_SUB
        upsc=[m for m in models if cat(m)=="超分"]
        for i,m in enumerate(upsc):
            ws6.cell(row=ur+1+i,column=2,value=m)
            ws6.cell(row=ur+1+i,column=4,value=model[m]["积分"]).number_format="#,##0"
            ws6.cell(row=ur+1+i,column=5,value=model[m]["次数"]).number_format="#,##0"
            ws6.cell(row=ur+1+i,column=9,value="有做超分").font=Font(color=ORANGE)
        note(ws6,ur+1+len(upsc)+1,"口径：单次积分 = 消耗积分 ÷ 调用次数；配置档位数 = 该模型出现过几种不同单次积分（≈不同分辨率/时长配置；源数据无分辨率字段）")
        ws6.conditional_formatting.add(f"D{f0}:D{last}",DataBarRule(start_type="num",start_value=0,end_type="max",color=PRIMARY,showValue=True))
        for j,w_ in enumerate([6,22,11,12,13,12,15,14,12,9],1): ws6.column_dimensions[get_column_letter(j)].width=w_
        ws6.freeze_panes="A3"

    # ---- 充值对账 ----
    ws7=wb.create_sheet("充值对账"); ttl(ws7,"充值 / 赠送 对账","F")
    for j,h in enumerate(["时间","积分类型","变动积分","币种"],1): ws7.cell(row=2,column=j,value=h)
    hrow(ws7,2,4); f0=3
    for i,rr_ in enumerate(sorted(D["rech"],key=lambda r:r[F["time"]])):
        rr=f0+i
        ws7.cell(row=rr,column=1,value=rr_[F["time"]])
        ws7.cell(row=rr,column=2,value=(rr_[F["ptype"]] if F["ptype"] else "充值"))
        ws7.cell(row=rr,column=3,value=D["pv"](rr_)).number_format="#,##0"
        ws7.cell(row=rr,column=4,value="-").alignment=AC
        for j in range(1,5): ws7.cell(row=rr,column=j).border=Border(bottom=thin)
    r=f0+len(D["rech"])+1
    rows_sum=[("充值总额（积分）",POOL,"#,##0",False),("　实充",PAID,"#,##0",False),("　赠送",POOL-PAID,"#,##0",False),
              ("累计消耗（积分）",TOT_JI,"#,##0",False),("消耗进度",(f"{TOT_JI/POOL*100:.1f}%" if POOL else "-"),None,True),
              ("实充积分已消耗",PAID_USED,"#,##0",True),("当前剩余（积分）",BALANCE,"#,##0",False),
              ("按近7天日均预计可支撑",f"{left:.1f} 天",None,True)]
    for lab,val,fmt,dg in rows_sum:
        a=ws7.cell(row=r,column=1,value=lab); a.border=Border(bottom=thin)
        b=ws7.cell(row=r,column=2,value=val); b.border=Border(bottom=thin); b.alignment=AR
        if fmt: b.number_format=fmt
        if dg: a.font=Font(bold=True,color=DANGER); b.font=Font(bold=True,color=DANGER)
        r+=1
    for j,w_ in enumerate([32,16,14,10],1): ws7.column_dimensions[get_column_letter(j)].width=w_

    # ---- 原始流水 ----
    ws8=wb.create_sheet("原始流水")
    cols=["序号","时间","日期",pname]+(["AI模型","类别"] if has_model else [])+["积分类型","变动积分","金额(元)","余额"]
    for j,h in enumerate(cols,1): ws8.cell(row=1,column=j,value=h)
    hrow(ws8,1,len(cols))
    for i,rr_ in enumerate(sorted(D["cons"],key=lambda r:r[F["time"]],reverse=True)):
        rr=i+2
        base=[i+1,rr_[F["time"]],rr_["_d"],disp(rr_["_p"])]
        if has_model: base+=[rr_[F["model"]],cat(rr_[F["model"]])]
        base+=[(rr_[F["ptype"]] if F["ptype"] else ""),D["pv"](rr_),round(D["amt"](rr_),2) if has_amt else "-"]
        for j,v in enumerate(base,1): ws8.cell(row=rr,column=j,value=v)
        ws8.cell(row=rr,column=1).alignment=AC
        if i%2==1:
            for j in range(1,len(cols)+1): ws8.cell(row=rr,column=j).fill=FZ
    for j in range(1,len(cols)+1): ws8.column_dimensions[get_column_letter(j)].width=[6,19,12,14,22,11,18,12,11][j-1] if j<=9 else 11
    ws8.freeze_panes="A2"; ws8.auto_filter.ref=f"A1:{get_column_letter(len(cols))}{TOT_CNT+1}"

    # ---- 按月汇总（附于按周汇总表底部，作月度图数据源） ----
    month=defaultdict(lambda:{"积分":0,"次数":0,"日":set()})
    for d,a in day.items():
        mm=month[d[:7]]; mm["积分"]+=a["积分"]; mm["次数"]+=a["次数"]; mm["日"].add(d)
    mos=sorted(month)
    mr0=f0+len(weeks)+2
    ws3.cell(row=mr0,column=1,value="按月汇总").font=F_SUB
    for j,h in enumerate(["月份","消耗积分","调用次数(次)","活跃天数"],1): ws3.cell(row=mr0+1,column=j,value=h)
    hrow(ws3,mr0+1,4)
    for i,mo in enumerate(mos):
        rr=mr0+2+i; mm=month[mo]
        ws3.cell(row=rr,column=1,value=mo).alignment=AC
        ws3.cell(row=rr,column=2,value=mm["积分"]).number_format="#,##0"
        ws3.cell(row=rr,column=3,value=mm["次数"]).number_format="#,##0"
        ws3.cell(row=rr,column=4,value=len(mm["日"])).alignment=AC
        for j in range(1,5): ws3.cell(row=rr,column=j).border=Border(bottom=thin)
    mlast=mr0+1+len(mos)

    # ---- 图表（放总览：日柱 + 周折线 + 月柱） ----
    ws_top=wb["总览"]
    ch=BarChart(); ch.type="col"; ch.title="按天消耗积分趋势"; ch.style=2
    ch.y_axis.title="积分"; ch.x_axis.title="日期"
    ch.add_data(Reference(wb["按天汇总"],min_col=4,min_row=2,max_row=1+len(days)),titles_from_data=True)
    ch.set_categories(Reference(wb["按天汇总"],min_col=1,min_row=3,max_row=2+len(days)))
    for s in ch.series: s.graphicalProperties.solidFill=PRIMARY; s.graphicalProperties.line.solidFill=PRIMARY
    ch.width=19; ch.height=8.5; ch.gapWidth=40
    ws_top.add_chart(ch,"H3")
    ch2=LineChart(); ch2.title="按周消耗积分趋势"; ch2.style=2
    ch2.y_axis.title="积分"; ch2.x_axis.title="周（周一起始日）"
    ch2.add_data(Reference(wb["按周汇总"],min_col=3,min_row=2,max_row=1+len(weeks)),titles_from_data=True)
    ch2.set_categories(Reference(wb["按周汇总"],min_col=1,min_row=3,max_row=2+len(weeks)))
    for s in ch2.series:
        s.graphicalProperties.line.solidFill=TEAL; s.graphicalProperties.line.width=28000
        s.marker=Marker(symbol="circle",size=6)
    ch2.width=19; ch2.height=8
    ws_top.add_chart(ch2,"H22")
    ch3=BarChart(); ch3.type="col"; ch3.title="按月消耗积分"; ch3.style=2
    ch3.y_axis.title="积分"; ch3.x_axis.title="月份"
    ch3.add_data(Reference(ws3,min_col=2,min_row=mr0+1,max_row=mlast),titles_from_data=True)
    ch3.set_categories(Reference(ws3,min_col=1,min_row=mr0+2,max_row=mlast))
    for s in ch3.series: s.graphicalProperties.solidFill=ORANGE; s.graphicalProperties.line.solidFill=ORANGE
    labels(ch3,900)
    ch3.width=19; ch3.height=8
    ws_top.add_chart(ch3,"H41")

    wb.properties.title="积分消耗分析（客户视角）"
    span_tag=f"{days[0].replace('-','')[2:]}-{days[-1].replace('-','')[2:]}"
    gone="耗尽" if (POOL and TOT_JI/POOL>=0.95) else "进行中"
    out_path=os.path.join(out,f"积分消耗分析_客户视角_{ENT[:12]}_{span_tag}消耗{gone}.xlsx")
    wb.save(out_path); return out_path

# ============ 视角二：PIXMAX 业务 ============
def build_business(D,out):
    F=D["F"]; people=D["people"]; person=D["person"]
    TOT_JI=D["TOT_JI"]; TOT_AMT=D["TOT_AMT"]; TOT_CNT=D["TOT_CNT"]
    days=D["days"]; day=D["day"]; weeks=D["weeks"]; week=D["week"]
    models=D["models"]; model=D["model"]; BALANCE=D["BALANCE"]
    PAID=D["PAID"]; POOL=D["POOL"]; PAID_USED=D["PAID_USED"]; r7avg=D["r7avg"]; left=D["left"]
    has_amt=D["has_amt"]; has_model=D["has_model"]; ENT=D["company"]
    span=D["span"]; used_days=len(days); gap=span-used_days
    avg_all=D["avg_all"]; eta=D["eta"]; d0=D["d0"]; d1=D["d1"]
    mid=days[0]
    # 前期/近期
    cut=date(*map(int,days[len(days)//2].split("-")))
    early=[r for r in D["cons"] if date(*map(int,r["_d"].split("-")))<=cut]
    late=[r for r in D["cons"] if date(*map(int,r["_d"].split("-")))>cut]
    ey=len({r["_d"] for r in early}); ly=len({r["_d"] for r in late})
    EJ=sum(D["pv"](r) for r in early); LJ=sum(D["pv"](r) for r in late)
    ea=EJ/ey if ey else 0; la=LJ/ly if ly else 0; accel=la/ea if ea else 0
    gap_days=[]
    cur=d0
    while cur<=d1:
        if cur.isoformat() not in day: gap_days.append(cur.isoformat())
        cur+=timedelta(days=1)
    conc1=person[people[0]]["积分"]/TOT_JI
    top3=sum(person[k]["积分"] for k in people[:3])/TOT_JI
    video_share=sum(model[m]["积分"] for m in models if cat(m)=="视频生成")/TOT_JI if has_model else 1

    wb=Workbook()
    ws=wb.active; ws.title="客户经营速览"
    ttl(ws,f"PIXMAX 客户经营速览 · {ENT}","F")
    r=3
    c=ws.cell(row=r,column=1,value=f"核心结论：该客户预计 {eta} 前后耗尽积分（剩 {left:.1f} 天），建议尽快触达催充；消耗处于加速状态（近期日均是前期的 {accel:.1f} 倍）。")
    c.font=Font(bold=True,size=12,color=DANGER); c.fill=FW; c.alignment=AL; ws.merge_cells(f"A{r}:F{r}")
    r+=2
    ws.cell(row=r,column=1,value="① 客户健康度").font=F_SUB; r+=1
    for j,h in enumerate(["指标","数值","解读"],1): ws.cell(row=r,column=j,value=h)
    ws.cell(row=r,column=3,value="解读"); ws.merge_cells(f"C{r}:F{r}")
    hrow(ws,r,6)
    health=[("充值总额（积分）",f"{POOL:,}",f"实充 {PAID:,} + 赠送 {POOL-PAID:,}"),
            ("已消耗（积分）",f"{TOT_JI:,}",f"占充值总额 {TOT_JI/POOL*100:.1f}%" if POOL else "-"),
            ("剩余（积分）",f"{BALANCE:,}","实充部分已用尽" if PAID_USED>=PAID and PAID else "仍有实充余额"),
            ("消耗天数",f"{span} 天",f"{days[0]} ~ {days[-1]}，实际有消耗 {used_days} 天"),
            ("日均消耗（全程）",f"{round(avg_all):,}","按自然日摊"),
            ("日均消耗（近7天）",f"{round(r7avg):,}",f"较全程高 {r7avg/avg_all*100-100:.0f}%"),
            ("预计耗尽日",eta,f"按近7天速度，剩 {left:.1f} 天"),
            ("零消耗天数",f"{gap} 天",f"共 {span} 天中 {gap} 天无调用")]
    for i,(a,b,cc) in enumerate(health):
        rr=r+1+i
        ws.cell(row=rr,column=1,value=a).font=F_B
        y=ws.cell(row=rr,column=2,value=b); y.alignment=AR; y.font=F_B
        z=ws.cell(row=rr,column=3,value=cc); z.alignment=AL; ws.merge_cells(f"C{rr}:F{rr}")
        if a in ("预计耗尽日","日均消耗（近7天）","剩余（积分）"): y.font=Font(bold=True,color=DANGER)
        for j in range(1,7): ws.cell(row=rr,column=j).border=Border(bottom=thin)
    r=r+1+len(health)+1
    ws.cell(row=r,column=1,value="② 消耗节奏：前期 vs 近期").font=F_SUB; r+=1
    for j,h in enumerate(["阶段","日期区间","天数","消耗积分","日均消耗"],1): ws.cell(row=r,column=j,value=h)
    hrow(ws,r,5)
    for i,(a,b,nd,jj,av) in enumerate([("前期",f"{days[0]} ~ {cut.isoformat()}",ey,EJ,round(ea)),("近期",(cut+timedelta(days=1)).isoformat()+f" ~ {days[-1]}",ly,LJ,round(la))]):
        rr=r+1+i
        ws.cell(row=rr,column=1,value=a).font=F_B
        ws.cell(row=rr,column=2,value=b).alignment=AC
        ws.cell(row=rr,column=3,value=nd).alignment=AC
        ws.cell(row=rr,column=4,value=jj).number_format="#,##0"
        ws.cell(row=rr,column=5,value=av).number_format="#,##0"
        for j in range(1,6): ws.cell(row=rr,column=j).border=Border(bottom=thin)
    rr=r+3
    c=ws.cell(row=rr,column=1,value=f"→ 近期日均消耗是前期的 {accel:.1f} 倍。"); c.font=Font(bold=True,color=GREEN); ws.merge_cells(f"A{rr}:F{rr}")
    r=rr+2
    ws.cell(row=r,column=1,value="③ 建议动作（PIXMAX 侧）").font=F_SUB; r+=1
    acts=[f"1. 尽快触达催充：预计 {eta} 耗尽，只剩 {left:.1f} 天。",
          f"2. 顺势加售：用量在加速（{accel:.1f} 倍），可推大额套餐/年包。" ,
          f"3. 关注断档：{gap} 天零消耗（{'、'.join(g[5:] for g in gap_days)}），确认节假日或项目间歇。" ,
          f"4. 结构健康：消耗分散在 {len(people)} 个账号，Top1 仅占 {conc1*100:.1f}%，无单点依赖。"]
    for i,a in enumerate(acts):
        c=ws.cell(row=r+i,column=1,value=a); c.alignment=AL; ws.merge_cells(f"A{r+i}:F{r+i}")
    for j,w_ in enumerate([24,22,10,16,16,30],1): ws.column_dimensions[get_column_letter(j)].width=w_

    ws2=wb.create_sheet("消耗节奏与趋势"); ttl(ws2,"消耗节奏与趋势","G")
    for j,h in enumerate(["日期","星期","消耗积分","调用次数(次)","活跃账号(个)","较上日变化","状态"],1): ws2.cell(row=2,column=j,value=h)
    hrow(ws2,2,7); f0=3; prev=None
    cur=d0; seq=[]
    while cur<=d1: seq.append(cur.isoformat()); cur+=timedelta(days=1)
    for i,d in enumerate(seq):
        rr=f0+i; a=day.get(d)
        ws2.cell(row=rr,column=1,value=d).alignment=AC
        ws2.cell(row=rr,column=2,value=wd(d)).alignment=AC
        if a:
            ji=a["积分"]
            ws2.cell(row=rr,column=3,value=ji).number_format="#,##0"
            ws2.cell(row=rr,column=4,value=a["次数"]).number_format="#,##0"
            ws2.cell(row=rr,column=5,value=len(a["用户"])).alignment=AC
            if prev is not None:
                cc=ws2.cell(row=rr,column=6,value=ji-prev); cc.number_format="+#,##0;-#,##0"; cc.alignment=AR
                cc.font=Font(color=DANGER if ji-prev>0 else GREEN)
            ws2.cell(row=rr,column=7,value="消耗").alignment=AC
            prev=ji
        else:
            ws2.cell(row=rr,column=3,value=0).number_format="#,##0"
            cc=ws2.cell(row=rr,column=7,value="⚠ 断档"); cc.alignment=AC; cc.font=Font(color=ORANGE)
            for j in range(1,7): ws2.cell(row=rr,column=j).fill=FZ
        for j in range(1,8): ws2.cell(row=rr,column=j).border=Border(bottom=thin)
    last=f0+len(seq)-1; sr=last+1
    ws2.cell(row=sr,column=1,value="合计").alignment=AC
    ws2.cell(row=sr,column=3,value=TOT_JI).number_format="#,##0"
    ws2.cell(row=sr,column=4,value=TOT_CNT).number_format="#,##0"
    for j in range(1,8):
        c=ws2.cell(row=sr,column=j); c.font=F_SUMF; c.fill=FS; c.alignment=AC
    ws2.conditional_formatting.add(f"C{f0}:C{last}",DataBarRule(start_type="num",start_value=0,end_type="max",color=PRIMARY,showValue=True))
    for j,w_ in enumerate([12,8,12,14,12,14,10],1): ws2.column_dimensions[get_column_letter(j)].width=w_
    ws2.freeze_panes="A3"
    note(ws2,sr+2,f"断档日（灰底）= 当天无任何调用；共 {gap} 天。可用于判断客户停工/流失风险。")

    ws3=wb.create_sheet("账号结构"); ttl(ws3,"子账号结构与依赖风险","G")
    for j,h in enumerate(["排名","账号","消耗积分","调用次数(次)","活跃天数","日均消耗","占总比"],1): ws3.cell(row=2,column=j,value=h)
    hrow(ws3,2,7); f0=3
    for i,s in enumerate(people):
        rr=f0+i; p=person[s]
        dset={r["_d"] for r in D["cons"] if r["_p"]==s}
        ws3.cell(row=rr,column=1,value=i+1).alignment=AC
        ws3.cell(row=rr,column=2,value=s)
        ws3.cell(row=rr,column=3,value=p["积分"]).number_format="#,##0"
        ws3.cell(row=rr,column=4,value=p["次数"]).number_format="#,##0"
        ws3.cell(row=rr,column=5,value=len(dset)).alignment=AC
        ws3.cell(row=rr,column=6,value=round(p["积分"]/len(dset)) if dset else 0).number_format="#,##0"
        ws3.cell(row=rr,column=7,value=f"{p['积分']/TOT_JI*100:.1f}%").alignment=AC
        for j in range(1,8): ws3.cell(row=rr,column=j).border=Border(bottom=thin)
    last=f0+len(people)-1; sr=last+1
    ws3.cell(row=sr,column=2,value="合计")
    ws3.cell(row=sr,column=3,value=TOT_JI).number_format="#,##0"
    ws3.cell(row=sr,column=4,value=TOT_CNT).number_format="#,##0"
    ws3.cell(row=sr,column=7,value="100.0%").alignment=AC
    for j in range(1,8):
        c=ws3.cell(row=sr,column=j); c.font=F_SUMF; c.fill=FS; c.alignment=AC
    ws3.conditional_formatting.add(f"C{f0}:C{last}",DataBarRule(start_type="num",start_value=0,end_type="max",color=PRIMARY,showValue=True))
    for j,w_ in enumerate([6,14,12,14,10,12,10],1): ws3.column_dimensions[get_column_letter(j)].width=w_
    note(ws3,sr+2,f"集中度：Top1 占 {conc1*100:.1f}%，Top3 合计 {top3*100:.1f}%。"+("分散度健康（无单点依赖）。" if conc1<0.3 else "集中度偏高（存在单点依赖风险）。"))

    if has_model:
        ws4=wb.create_sheet("模型偏好"); ttl(ws4,"客户模型偏好（产品/销售参考）","H")
        for j,h in enumerate(["排名","AI模型","类别","消耗积分","调用次数(次)","单次积分","占总比","调用账号(个)"],1): ws4.cell(row=2,column=j,value=h)
        hrow(ws4,2,8); f0=3
        for i,m in enumerate(models):
            rr=f0+i; d=model[m]
            ws4.cell(row=rr,column=1,value=i+1).alignment=AC
            ws4.cell(row=rr,column=2,value=m)
            ws4.cell(row=rr,column=3,value=cat(m)).alignment=AC
            ws4.cell(row=rr,column=4,value=d["积分"]).number_format="#,##0"
            ws4.cell(row=rr,column=5,value=d["次数"]).number_format="#,##0"
            ws4.cell(row=rr,column=6,value=round(d["积分"]/d["次数"],1)).number_format="0.0"
            ws4.cell(row=rr,column=7,value=f"{d['积分']/TOT_JI*100:.1f}%").alignment=AC
            ws4.cell(row=rr,column=8,value=len(d["用户"])).alignment=AC
            for j in range(1,9): ws4.cell(row=rr,column=j).border=Border(bottom=thin)
        last=f0+len(models)-1; sr=last+1
        ws4.cell(row=sr,column=2,value="合计")
        ws4.cell(row=sr,column=4,value=TOT_JI).number_format="#,##0"
        ws4.cell(row=sr,column=5,value=TOT_CNT).number_format="#,##0"
        ws4.cell(row=sr,column=7,value="100.0%").alignment=AC
        for j in range(1,9):
            c=ws4.cell(row=sr,column=j); c.font=F_SUMF; c.fill=FS; c.alignment=AC
        ws4.conditional_formatting.add(f"D{f0}:D{last}",DataBarRule(start_type="num",start_value=0,end_type="max",color=PRIMARY,showValue=True))
        for j,w_ in enumerate([6,22,11,12,14,10,9,13],1): ws4.column_dimensions[get_column_letter(j)].width=w_

    ws5=wb.create_sheet("充值与续费"); ttl(ws5,"充值记录与续费预测","F")
    for j,h in enumerate(["时间","积分类型","变动积分","币种"],1): ws5.cell(row=2,column=j,value=h)
    hrow(ws5,2,4); f0=3
    for i,rr_ in enumerate(sorted(D["rech"],key=lambda r:r[F["time"]])):
        rr=f0+i
        ws5.cell(row=rr,column=1,value=rr_[F["time"]])
        ws5.cell(row=rr,column=2,value=(rr_[F["ptype"]] if F["ptype"] else "充值"))
        ws5.cell(row=rr,column=3,value=D["pv"](rr_)).number_format="#,##0"
        ws5.cell(row=rr,column=4,value="-").alignment=AC
        for j in range(1,5): ws5.cell(row=rr,column=j).border=Border(bottom=thin)
    r=f0+len(D["rech"])+1
    for lab,val,fmt,dg in [("充值总额（积分）",POOL,"#,##0",False),("　实充",PAID,"#,##0",False),("　赠送",POOL-PAID,"#,##0",False),
        ("累计消耗（积分）",TOT_JI,"#,##0",False),("消耗进度",(f"{TOT_JI/POOL*100:.1f}%" if POOL else "-"),None,True),
        ("实充积分已消耗",PAID_USED,"#,##0",True),("当前剩余（积分）",BALANCE,"#,##0",False),
        ("预计耗尽日",eta,None,True),("从充值到耗尽的周期",f"约 {span+round(left)} 天",None,False)]:
        a=ws5.cell(row=r,column=1,value=lab); a.border=Border(bottom=thin)
        b=ws5.cell(row=r,column=2,value=val); b.border=Border(bottom=thin); b.alignment=AR
        if fmt: b.number_format=fmt
        if dg: a.font=Font(bold=True,color=DANGER); b.font=Font(bold=True,color=DANGER)
        r+=1
    note(ws5,r+1,"续费预测：预计耗尽日 = 最后消耗日 + 剩余积分 ÷ 近7天日均消耗。")
    for j,w_ in enumerate([28,16,14,10],1): ws5.column_dimensions[get_column_letter(j)].width=w_

    # ---- 周度/月度汇总（趋势数据源，写在逐日表下方） ----
    month=defaultdict(lambda:{"积分":0,"次数":0,"日":set()})
    for d,a in day.items():
        mm=month[d[:7]]; mm["积分"]+=a["积分"]; mm["次数"]+=a["次数"]; mm["日"].add(d)
    mos=sorted(month)
    wr0=len(seq)+7
    ws2.cell(row=wr0,column=1,value="按周汇总").font=F_SUB
    for j,h in enumerate(["周起始","消耗积分","调用次数(次)","活跃天数"],1): ws2.cell(row=wr0+1,column=j,value=h)
    hrow(ws2,wr0+1,4)
    for i,wk in enumerate(weeks):
        rr=wr0+2+i; w=week[wk]
        ws2.cell(row=rr,column=1,value=wk[5:]).alignment=AC
        ws2.cell(row=rr,column=2,value=w["积分"]).number_format="#,##0"
        ws2.cell(row=rr,column=3,value=w["次数"]).number_format="#,##0"
        ws2.cell(row=rr,column=4,value=len(w["日"])).alignment=AC
        for j in range(1,5): ws2.cell(row=rr,column=j).border=Border(bottom=thin)
    wlast=wr0+1+len(weeks)
    mr0=wlast+3
    ws2.cell(row=mr0,column=1,value="按月汇总").font=F_SUB
    for j,h in enumerate(["月份","消耗积分","调用次数(次)","活跃天数"],1): ws2.cell(row=mr0+1,column=j,value=h)
    hrow(ws2,mr0+1,4)
    for i,mo in enumerate(mos):
        rr=mr0+2+i; mm=month[mo]
        ws2.cell(row=rr,column=1,value=mo).alignment=AC
        ws2.cell(row=rr,column=2,value=mm["积分"]).number_format="#,##0"
        ws2.cell(row=rr,column=3,value=mm["次数"]).number_format="#,##0"
        ws2.cell(row=rr,column=4,value=len(mm["日"])).alignment=AC
        for j in range(1,5): ws2.cell(row=rr,column=j).border=Border(bottom=thin)
    mlast=mr0+1+len(mos)

    # ---- 图表：日柱（无标签）+ 周折线 + 月柱 ----
    ws_top=wb["客户经营速览"]
    ch=BarChart(); ch.type="col"; ch.title="按天消耗积分（含断档日）"; ch.style=2
    ch.y_axis.title="积分"; ch.x_axis.title="日期"
    ch.add_data(Reference(ws2,min_col=3,min_row=2,max_row=1+len(seq)),titles_from_data=True)
    ch.set_categories(Reference(ws2,min_col=1,min_row=3,max_row=2+len(seq)))
    for s in ch.series: s.graphicalProperties.solidFill=PRIMARY; s.graphicalProperties.line.solidFill=PRIMARY
    ch.legend=None; ch.y_axis.majorGridlines=ChartLines()
    ch.width=20; ch.height=8.5; ch.gapWidth=30
    ws_top.add_chart(ch,"H3")
    chw=LineChart(); chw.title="按周消耗积分趋势"; chw.style=2
    chw.y_axis.title="积分"; chw.x_axis.title="周（周一起始日）"
    chw.add_data(Reference(ws2,min_col=2,min_row=wr0+1,max_row=wlast),titles_from_data=True)
    chw.set_categories(Reference(ws2,min_col=1,min_row=wr0+2,max_row=wlast))
    for s in chw.series:
        s.graphicalProperties.line.solidFill=TEAL; s.graphicalProperties.line.width=28000
        s.marker=Marker(symbol="circle",size=6)
    chw.legend=None; chw.y_axis.majorGridlines=ChartLines()
    chw.width=20; chw.height=8
    ws_top.add_chart(chw,"H21")
    chm=BarChart(); chm.type="col"; chm.title="按月消耗积分"; chm.style=2
    chm.y_axis.title="积分"; chm.x_axis.title="月份"
    chm.add_data(Reference(ws2,min_col=2,min_row=mr0+1,max_row=mlast),titles_from_data=True)
    chm.set_categories(Reference(ws2,min_col=1,min_row=mr0+2,max_row=mlast))
    for s in chm.series: s.graphicalProperties.solidFill=ORANGE; s.graphicalProperties.line.solidFill=ORANGE
    labels(chm,900)
    chm.width=20; chm.height=8
    ws_top.add_chart(chm,"H38")

    wb.properties.title="PIXMAX 客户经营分析"
    out_path=os.path.join(out,f"积分消耗分析_业务视角_{ENT[:12]}.xlsx")
    wb.save(out_path); return out_path

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("csv")
    ap.add_argument("--view",choices=["client","business","both"],default="client")
    ap.add_argument("--out",default=os.getcwd())
    a=ap.parse_args()
    os.makedirs(a.out,exist_ok=True)
    D=load(a.csv)
    print(f"[字段识别] 人={D['F']['person']} 模型={D['F']['model']} 金额={D['F']['amount']} 积分类型={D['F']['ptype']}")
    print(f"[概览] 客户={D['company']} 记录={D['TOT_CNT']} 消耗={D['TOT_JI']:,} 会 {len(D['people'])} 人 {len(D['days'])} 天")
    if a.view in ("client","both"): print("[输出] 客户视角:",build_client(D,a.out))
    if a.view in ("business","both"): print("[输出] 业务视角:",build_business(D,a.out))

if __name__=="__main__":
    main()
