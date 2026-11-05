import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference

VER="1.0"; REL="November 2026"
ALL="T1,T2,T3"; T1="T1"; T23="T2,T3"; T2="T2"
P="Pending SEBON directive"
D={
"D01":"Governance and Policy Implementation",
"D02":"Information Security Management",
"D03":"Access, Authentication and Digital Signature",
"D04":"API Security",
"D05":"Third-Party and Outsourcing Risk",
"D06":"Business Continuity, DR and Backup",
"D07":"Security Testing and IT Audit",
"D08":"Security Monitoring and Incident Response",
"D09":"Data Management and Localization",
"D10":"AI/ML and Algorithmic Trading",
}
Q=[
("D01","Management has approved an implementation plan, with timelines, for the requirements of SEBON IT Policy 2083 that apply to the organization.","10.1.2",ALL,""),
("D01","Responsibility for IT policy implementation and monitoring is assigned to a named function or committee with periodic reporting to senior management.","10.1.2, 10.1.3",ALL,""),
("D01","A Business Process Re-engineering (BPR) assessment is planned or performed every three years, covering digitization of licensing, regulatory reporting and surveillance-related processes.","10.1.2",ALL,""),
("D01","The organization tracks and complies with the minimum IT Standard issued by SEBON for licensed entities.","8.1.6",ALL,P),
("D01","The organizational structure includes a defined IT/IS function with adequate staffing for IT, resources and personnel management.","9.5.1",T1,""),
("D02","An Information Security Management System (ISMS) based on ISO/IEC 27001, or an equivalent standard approved by SEBON, is implemented.","9.1.5",T1,""),
("D02","The ISMS is ISO/IEC 27001 certified.","8.3.1",T1,""),
("D02","The organization complies with the minimum cyber security baseline prescribed by SEBON for market participants.","9.1.5",T23,P),
("D02","Information security roles, responsibilities and a cyber security awareness program are in place for staff.","11.1.3",ALL,""),
("D03","Multi-Factor Authentication (MFA) is enforced before sensitive information is transmitted or accessed in regulatory, trading, depository and intermediary service systems.","9.2.2",ALL,""),
("D03","Regulatory documents, reports and disclosures submitted to SEBON are digitally signed.","9.2.5",ALL,""),
("D03","Digital signature certificates in use are issued by a CA licensed under the Controller of Certification Authorities (CCA) or a CA recognized by SEBON.","9.2.5",ALL,""),
("D03","Digital signatures are used, or planned, for investor-broker agreements and KYC processes.","9.2.5",T2,""),
("D04","APIs exposed to third parties enforce authentication and authorization controls.","9.2.4",ALL,""),
("D04","APIs exposed to third parties enforce request rate limiting.","9.2.4",ALL,""),
("D04","API traffic is encrypted using TLS 1.2 or higher, with weaker protocols disabled.","9.2.4",ALL,""),
("D04","API activity is logged, retained and reviewed.","9.2.4",ALL,""),
("D04","APIs comply with the API security standard prescribed by SEBON.","9.2.4",ALL,P),
("D05","Vendor due diligence is performed before onboarding third-party IT service providers, including cloud providers.","8.6.1",ALL,""),
("D05","An ongoing vendor risk management process exists (inventory, risk rating, periodic review) for third-party IT service providers.","8.6.1",ALL,""),
("D05","IT outsourcing contracts include a Service Level Agreement (SLA) with measurable service levels.","8.1.5, 9.1.4",ALL,""),
("D05","IT outsourcing contracts define data security, confidentiality and liability obligations of the vendor.","8.1.5, 9.1.4",ALL,""),
("D05","IT outsourcing contracts grant audit and inspection rights, including regulatory audit rights.","8.1.5, 9.1.4",ALL,""),
("D05","IT outsourcing contracts include a documented vendor exit strategy.","8.1.5, 9.1.4",ALL,""),
("D05","Cloud computing is formally evaluated as an option when procuring IT equipment and services.","8.6.1, 9.1.4",ALL,""),
("D06","A Business Continuity Plan (BCP) is documented and approved.","9.1.3",ALL,""),
("D06","A Disaster Recovery Plan (DRP) is documented and approved.","9.1.3",ALL,""),
("D06","BCP and DRP are tested at least annually (drill or simulation) and the test report is submitted to SEBON.","9.1.3",ALL,""),
("D06","Recovery Time Objective (RTO) and Recovery Point Objective (RPO) are clearly defined for critical systems.","9.1.1",T1,""),
("D06","The Disaster Recovery Site is located at least 25 km from the primary site.","9.1.1",T1,""),
("D06","A multi-layered backup strategy protects sensitive data and systems from ransomware, malware, cyber attack and accidental modification or deletion.","9.1.2",ALL,""),
("D06","Immutable snapshots are used as part of the backup strategy.","9.1.2",ALL,""),
("D06","Air-gapped backups are maintained.","9.1.2",ALL,""),
("D07","VAPT of critical systems is conducted at least annually.","8.3.1, 9.3.2",ALL,""),
("D07","VAPT reports are submitted to SEBON within three months of test completion.","9.3.2",ALL,""),
("D07","Critical vulnerabilities identified in VAPT are remediated within 30 days.","9.3.2",ALL,""),
("D07","IT/system audit is conducted annually.","9.3.1",T1,""),
("D07","IT/system audit is conducted at least every two years.","9.3.1",T23,""),
("D07","IT audit scope covers, at minimum, trading systems, clearing and settlement, cyber security controls, BCP/DRP and data management.","9.3.1",ALL,""),
("D07","IT audit reports are submitted to SEBON within three months.","9.3.1",ALL,""),
("D07","Critical audit findings are remediated within 30 days.","9.3.1",ALL,""),
("D07","Defined standards are applied for software development or acquisition, technical evaluation and pre-delivery inspection.","9.3.2",ALL,""),
("D07","IT auditor selection follows the System Audit Framework and auditor selection norms issued by SEBON.","9.3.1",ALL,P),
("D08","24x7 cyber threat monitoring is in place through an in-house Security Operations Centre (SOC) or a recognized Managed Security Service Provider (MSSP).","9.3.3",T1,""),
("D08","An Incident Response Plan (IRP) is prepared and submitted to SEBON.","9.3.4",ALL,""),
("D08","Cyber incidents directly affecting market operations are reported to SEBON on the same day.","9.3.4",ALL,""),
("D08","Other significant cyber incidents are reported to SEBON within 72 hours.","9.3.4",ALL,""),
("D08","A Root Cause Analysis (RCA) report is submitted to SEBON within 30 days of an incident.","9.3.4",ALL,""),
("D08","Incidents are reported through the integrated capital market cyber incident reporting portal.","8.3.2, 9.3.4",ALL,P),
("D09","Data is classified into SEBON categories: public data, sensitive business data and personal investor data.","9.6.2",ALL,P),
("D09","Securities market data is stored within Nepal (data localization).","8.6.2, 9.6.2",ALL,""),
("D09","Data centres used are IDMC or data centres officially listed with the Department of Information Technology, Government of Nepal.","8.6.2",ALL,""),
("D09","A Data Lifecycle Management Policy is implemented.","9.5.2",T1,""),
("D09","Electronic filing to SEBON is done through SEDFAR with digital signatures.","9.6.3",ALL,P),
("D10","Algorithmic trading systems are approved by the stock exchange before use (mark N/A if not used).","9.4.3",T2,""),
("D10","AI/ML systems used in trading, surveillance or investor services have mandatory human oversight (human-in-the-loop). Mark N/A if not used.","9.4.3",ALL,""),
("D10","Decisions made by AI systems are explainable and documented (Explainable AI). Mark N/A if not used.","9.4.3",ALL,""),
]
rows=[]
for i,(d,q,c,a,n) in enumerate(Q):
    rows.append((f"{d}-{sum(1 for x in Q[:i+1] if x[0]==d):02d}",d,D[d],q,c,a,n))

# CSV source for GitHub
with open("question-bank.csv","w",newline="",encoding="utf-8") as f:
    w=csv.writer(f); w.writerow(["ID","Domain Code","Domain","Requirement","SEBON IT Policy 2083 Clause","Applicable Tiers","Note"])
    w.writerows(rows)

F="Calibri"
NAVY="1F3864"; LIGHT="D9E1F2"; INPUT="FFF2CC"
thin=Side(style="thin",color="A6A6A6"); B=Border(left=thin,right=thin,top=thin,bottom=thin)
def hdr(c):
    c.font=Font(name=F,bold=True,color="FFFFFF",size=11); c.fill=PatternFill("solid",fgColor=NAVY)
    c.alignment=Alignment(wrap_text=True,vertical="center",horizontal="center"); c.border=B
def body(c,wrap=True,center=False,bold=False):
    c.font=Font(name=F,size=11,bold=bold); c.border=B
    c.alignment=Alignment(wrap_text=wrap,vertical="top",horizontal="center" if center else "left")

wb=Workbook()
# ---- About
ws=wb.active; ws.title="About"
ws.column_dimensions["A"].width=3; ws.column_dimensions["B"].width=110
lines=[
("SEBON IT Policy 2083 Readiness Self-Assessment",Font(name=F,bold=True,size=18,color=NAVY)),
(f"Version {VER} | {REL} | Topology Defense Pvt. Ltd., Kathmandu",Font(name=F,size=11,italic=True)),
("",None),
("Purpose",Font(name=F,bold=True,size=13,color=NAVY)),
("A free self-assessment for SEBON-licensed entities to gauge readiness against the Nepal Securities Board Information Technology Policy, 2083 (effective 2083-07-18). Each requirement is mapped to the policy clause it comes from.",None),
("",None),
("How to use",Font(name=F,bold=True,size=13,color=NAVY)),
("1. Go to the 'Profile' sheet and fill the yellow cells. Select your tier; requirements that do not apply to your tier are excluded from scoring automatically.",None),
("2. Go to the 'Assessment' sheet. For each applicable requirement, select a maturity score (0 to 4) or N/A in the yellow column. See 'Maturity Scale' for definitions.",None),
("3. Record the evidence reference (policy name, report, ticket, screenshot) and any remarks. Score only what you can evidence.",None),
("4. Review the 'Summary' sheet for domain scores and gap counts. Use 'Key Deadlines' as a quick reference for reporting timelines.",None),
("Example: Requirement D07-01 (annual VAPT). Last VAPT done in 2082, report submitted, critical items closed within 30 days, tracked in register -> Score 3; Evidence: 'VAPT Report Ref VR-2082-01, closure tracker'.",Font(name=F,size=11,italic=True)),
("",None),
("Tier classification (Clause 10.1.1)",Font(name=F,bold=True,size=13,color=NAVY)),
("Tier 1: Securities Board of Nepal, market operator (stock exchange), central depository service. Market Infrastructure Institutions (MIIs), including the clearing house, are treated as Tier 1 for the requirements in Clauses 9.1.5 and 9.3.",None),
("Tier 2: Securities brokers, securities dealers, merchant bankers.",None),
("Tier 3: Credit rating agencies, share registrars.",None),
("",None),
("Requirements marked 'Pending SEBON directive'",Font(name=F,bold=True,size=13,color=NAVY)),
("Some requirements depend on standards SEBON has stated it will issue (for example the minimum cyber security baseline, API security standard, Data Classification Framework, System Audit Framework, SEDFAR and the incident reporting portal). These are included so you can prepare, and will be updated when the directives are published.",None),
("",None),
("Disclaimer",Font(name=F,bold=True,size=13,color=NAVY)),
("This tool is an independent readiness aid. It is not issued, reviewed or endorsed by SEBON, and a score from this tool does not represent compliance, certification or audit assurance. Requirements are paraphrased from the policy; the official Nepali text of SEBON IT Policy 2083 prevails in case of any difference. Always refer to the latest directives issued by SEBON.",None),
("",None),
("Contact",Font(name=F,bold=True,size=13,color=NAVY)),
("Topology Defense Pvt. Ltd. | For a guided gap assessment, roadmap or IT audit support, contact Topology Defense. [Add email / website here]",None),
("License: CC BY 4.0. You may share and adapt with attribution to Topology Defense Pvt. Ltd.",None),
]
for i,(t,fnt) in enumerate(lines,start=1):
    c=ws.cell(row=i,column=2,value=t); c.font=fnt or Font(name=F,size=11)
    c.alignment=Alignment(wrap_text=True,vertical="top")
ws.sheet_view.showGridLines=False

# ---- Profile
pf=wb.create_sheet("Profile")
pf.column_dimensions["A"].width=3; pf.column_dimensions["B"].width=32; pf.column_dimensions["C"].width=45
pf["B2"]="Organization Profile"; pf["B2"].font=Font(name=F,bold=True,size=16,color=NAVY)
items=["Organization name","Entity type","Tier","Assessment date","Assessed by (name, designation)","Reviewed by"]
for i,it in enumerate(items):
    r=4+i; body(pf.cell(row=r,column=2,value=it),bold=True)
    c=pf.cell(row=r,column=3); body(c); c.fill=PatternFill("solid",fgColor=INPUT)
pf["C6"]="Tier 2"
pf["C7"].number_format="yyyy-mm-dd"
dv=DataValidation(type="list",formula1='"Tier 1,Tier 2,Tier 3"',allow_blank=False); pf.add_data_validation(dv); dv.add("C6")
dv2=DataValidation(type="list",formula1='"Stock Exchange,Central Depository,Clearing House,Securities Broker,Securities Dealer,Merchant Banker,Credit Rating Agency,Share Registrar,Other"',allow_blank=True)
pf.add_data_validation(dv2); dv2.add("C5")
pf["B11"]="Yellow cells are inputs. Tier drives which requirements apply on the Assessment sheet."
pf["B11"].font=Font(name=F,italic=True,size=10)
pf.sheet_view.showGridLines=False

# ---- Maturity Scale
ms=wb.create_sheet("Maturity Scale")
ms.column_dimensions["A"].width=3; ms.column_dimensions["B"].width=10; ms.column_dimensions["C"].width=24; ms.column_dimensions["D"].width=80
ms["B2"]="Maturity Scale"; ms["B2"].font=Font(name=F,bold=True,size=16,color=NAVY)
for j,h in enumerate(["Score","Level","Description"]): hdr(ms.cell(row=4,column=2+j,value=h))
scale=[(0,"Not in place","Requirement is not addressed. No process, control or documentation exists."),
(1,"Initial","Done informally or inconsistently. Depends on individuals. Little or no documentation."),
(2,"Defined","Documented and approved policy, procedure or plan exists, but implementation is partial or not yet consistent."),
(3,"Implemented","Implemented and operating consistently across scope. Evidence is available on request."),
(4,"Managed","Implemented, monitored and periodically reviewed. Metrics or test results exist and improvements are tracked."),
("N/A","Not applicable","Requirement does not apply to the organization (for example, no APIs exposed or no algorithmic trading). Record the reason in Remarks.")]
for i,(s,l,d) in enumerate(scale):
    r=5+i; body(ms.cell(row=r,column=2,value=s),center=True); body(ms.cell(row=r,column=3,value=l),bold=True); body(ms.cell(row=r,column=4,value=d))
ms["B12"]="Target for regulatory readiness is generally 3 (Implemented) or above for every applicable requirement."
ms["B12"].font=Font(name=F,italic=True,size=10)
ms.sheet_view.showGridLines=False

# ---- Assessment
a=wb.create_sheet("Assessment")
heads=["ID","Domain","Requirement","Policy Clause","Applicable Tiers","Note","Applies to You","Maturity Score (0-4 / N/A)","Evidence Reference","Remarks / Action"]
widths=[10,26,62,13,11,16,10,13,28,30]
for j,(h,wd) in enumerate(zip(heads,widths),start=1):
    hdr(a.cell(row=1,column=j,value=h)); a.column_dimensions[a.cell(row=1,column=j).column_letter].width=wd
a.row_dimensions[1].height=36
for i,(rid,dc,dn,q,c,ap,n) in enumerate(rows):
    r=2+i
    for j,v in enumerate([rid,dn,q,c,ap,n],start=1):
        body(a.cell(row=r,column=j,value=v),center=(j in (1,4,5)))
    g=a.cell(row=r,column=7,value=f'=IF(ISNUMBER(SEARCH("T"&RIGHT(Profile!$C$6,1),E{r})),"Yes","No")'); body(g,center=True)
    for col in (8,9,10):
        cc=a.cell(row=r,column=col); body(cc,center=(col==8)); cc.fill=PatternFill("solid",fgColor=INPUT)
last=1+len(rows)
dv3=DataValidation(type="list",formula1='"0,1,2,3,4,N/A"',allow_blank=True,showErrorMessage=True,error="Select 0-4 or N/A")
a.add_data_validation(dv3); dv3.add(f"H2:H{last}")
grey=PatternFill("solid",fgColor="D9D9D9")
a.conditional_formatting.add(f"A2:J{last}",FormulaRule(formula=[f'$G2="No"'],fill=grey,font=Font(color="808080")))
for v,col in [(0,"F8CBAD"),(1,"F8CBAD"),(2,"FFE699"),(3,"C6EFCE"),(4,"A9D08E")]:
    a.conditional_formatting.add(f"H2:H{last}",FormulaRule(formula=[f'AND(ISNUMBER($H2),$H2={v})'],fill=PatternFill("solid",fgColor=col)))
a.freeze_panes="D2"; a.auto_filter.ref=f"A1:J{last}"

# ---- Summary
s=wb.create_sheet("Summary",1)
s.column_dimensions["A"].width=3
for col,wd in zip("BCDEFGH",[38,13,13,13,14,14,16]): s.column_dimensions[col].width=wd
s["B2"]="Readiness Summary"; s["B2"].font=Font(name=F,bold=True,size=16,color=NAVY)
s["B3"]='="Organization: "&IF(Profile!C4="","[not entered]",Profile!C4)&"   |   "&Profile!C6'; s["B3"].font=Font(name=F,size=11,italic=True)
sh=["Domain","Applicable","Scored","Not Applicable (N/A)","Average Maturity (0-4)","Readiness %","Gaps (score 0-1)"]
for j,h in enumerate(sh): hdr(s.cell(row=5,column=2+j,value=h))
s.row_dimensions[5].height=32
A="Assessment!"
for i,(dc,dn) in enumerate(D.items()):
    r=6+i
    body(s.cell(row=r,column=2,value=dn))
    s.cell(row=r,column=3,value=f'=COUNTIFS({A}$B$2:$B${last},$B{r},{A}$G$2:$G${last},"Yes")')
    s.cell(row=r,column=4,value=f'=COUNTIFS({A}$B$2:$B${last},$B{r},{A}$G$2:$G${last},"Yes",{A}$H$2:$H${last},">=0")')
    s.cell(row=r,column=5,value=f'=COUNTIFS({A}$B$2:$B${last},$B{r},{A}$G$2:$G${last},"Yes",{A}$H$2:$H${last},"N/A")')
    s.cell(row=r,column=6,value=f'=IFERROR(AVERAGEIFS({A}$H$2:$H${last},{A}$B$2:$B${last},$B{r},{A}$G$2:$G${last},"Yes"),"")')
    s.cell(row=r,column=7,value=f'=IF(F{r}="","",F{r}/4)')
    s.cell(row=r,column=8,value=f'=COUNTIFS({A}$B$2:$B${last},$B{r},{A}$G$2:$G${last},"Yes",{A}$H$2:$H${last},"<=1")')
    for col in range(3,9): body(s.cell(row=r,column=col),center=True)
    s.cell(row=r,column=6).number_format="0.0"; s.cell(row=r,column=7).number_format="0%"
tr=6+len(D)
body(s.cell(row=tr,column=2,value="Overall"),bold=True)
for col,L in zip(range(3,6),"CDE"):
    c=s.cell(row=tr,column=col,value=f"=SUM({L}6:{L}{tr-1})"); body(c,center=True,bold=True)
c=s.cell(row=tr,column=6,value=f'=IFERROR(AVERAGEIFS({A}$H$2:$H${last},{A}$G$2:$G${last},"Yes"),"")'); body(c,center=True,bold=True); c.number_format="0.0"
c=s.cell(row=tr,column=7,value=f'=IF(F{tr}="","",F{tr}/4)'); body(c,center=True,bold=True); c.number_format="0%"
c=s.cell(row=tr,column=8,value=f"=SUM(H6:H{tr-1})"); body(c,center=True,bold=True)
for col in range(2,9): s.cell(row=tr,column=col).fill=PatternFill("solid",fgColor=LIGHT)
s.cell(row=tr+1,column=2,value=f'=IF(D{tr}<C{tr}-E{tr},"Note: "&(C{tr}-E{tr}-D{tr})&" applicable requirement(s) not yet scored.","All applicable requirements scored.")').font=Font(name=F,italic=True,size=10,color="C00000")
s.cell(row=tr+2,column=2,value="Readiness % = average maturity / 4. Averages exclude N/A, unscored and non-applicable items. Gaps = applicable items scored 0 or 1.").font=Font(name=F,italic=True,size=10)
s.conditional_formatting.add(f"G6:G{tr}",CellIsRule(operator="lessThan",formula=["0.5"],fill=PatternFill("solid",fgColor="F8CBAD")))
s.conditional_formatting.add(f"G6:G{tr}",CellIsRule(operator="between",formula=["0.5","0.7499"],fill=PatternFill("solid",fgColor="FFE699")))
s.conditional_formatting.add(f"G6:G{tr}",CellIsRule(operator="greaterThanOrEqual",formula=["0.75"],fill=PatternFill("solid",fgColor="C6EFCE")))
ch=BarChart(); ch.type="bar"; ch.title="Average Maturity by Domain"; ch.style=10
ch.add_data(Reference(s,min_col=6,min_row=5,max_row=tr-1),titles_from_data=True)
ch.set_categories(Reference(s,min_col=2,min_row=6,max_row=tr-1))
ch.y_axis.scaling.min=0; ch.y_axis.scaling.max=4; ch.legend=None; ch.height=9; ch.width=22
ch.y_axis.delete=False; ch.x_axis.delete=False
s.add_chart(ch,f"B{tr+4}")
s.sheet_view.showGridLines=False

# ---- Key Deadlines
k=wb.create_sheet("Key Deadlines")
k.column_dimensions["A"].width=3
for col,wd in zip("BCDE",[52,30,14,22]): k.column_dimensions[col].width=wd
k["B2"]="Key Deadlines and Frequencies (SEBON IT Policy 2083)"; k["B2"].font=Font(name=F,bold=True,size=16,color=NAVY)
for j,h in enumerate(["Obligation","Deadline / Frequency","Clause","Applies to"]): hdr(k.cell(row=4,column=2+j,value=h))
dl=[("Report cyber incident directly affecting market operations","Same day","9.3.4","Market participants"),
("Report other significant cyber incidents","Within 72 hours","9.3.4","Market participants"),
("Submit Root Cause Analysis report","Within 30 days of incident","9.3.4","Market participants"),
("IT/system audit","Annual","9.3.1","Market infrastructure institutions"),
("IT/system audit","Every two years","9.3.1","Other market participants"),
("Submit IT audit report to SEBON","Within 3 months","9.3.1","All audited entities"),
("Remediate critical audit findings","Within 30 days","9.3.1","All audited entities"),
("VAPT of critical systems","Annual","9.3.2","All"),
("Submit VAPT report to SEBON","Within 3 months of test completion","9.3.2","All"),
("Remediate critical VAPT vulnerabilities","Within 30 days","9.3.2","All"),
("BCP and DRP drill/simulation, report to SEBON","Annual","9.1.3","Market infrastructure and professionals"),
("Business Process Re-engineering assessment","Every 3 years","10.1.2","Related parties"),
("DR site distance from primary site","At least 25 km","9.1.1","Market infrastructure (trading, clearing/settlement, CDS)")]
for i,row in enumerate(dl):
    for j,v in enumerate(row): body(k.cell(row=5+i,column=2+j,value=v),center=(j==2))
k.cell(row=6+len(dl),column=2,value="Paraphrased from the policy text. The official Nepali text prevails.").font=Font(name=F,italic=True,size=10)
k.sheet_view.showGridLines=False

for w in wb.worksheets:
    w.sheet_properties.tabColor=NAVY if w.title in ("About","Summary") else "8EA9DB"
    w.page_setup.orientation="landscape"; w.page_setup.fitToWidth=1; w.page_setup.fitToHeight=0
    w.sheet_properties.pageSetUpPr.fitToPage=True
a.print_title_rows="1:1"
wb.save("SEBON-IT-Policy-2083-Readiness-Assessment-v1.0.xlsx")
print(len(rows))
