import pandas as pd
import win32com.client

HUNDRED = '*' * 100

print(f'{HUNDRED}\nLayout Loader for the SAP R/3 standard t-code reports (FS10N, Customer and Vendor line-items, etc..)\n{HUNDRED}')
input('xlsx columns copied in clipboard ?? ... "Enter" to continue ...')
input('Only a single SAP R/3 session window allowed, close others, if opened ... "Enter to continue ...')
input('Ensure that SAP R/3 report results are already produced on the screen ... "Enter" to continue ...')

x = list(pd.read_clipboard(sep='\t'))  # ļ \t matches single tab.
order = [[en, t] for en, t in enumerate(x)]

session = win32com.client.GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[1]/btn[32]").press()

try:
    session.findById("wnd[1]/usr/btnAPP_FL_ALL").press()
    session.findById("wnd[1]/usr/btn%#AUTOTEXT002").press()
    for law in order:
        session.findById("wnd[1]/usr/btnB_SEARCH").press()  # open report layout fields
        session.findById("wnd[2]/usr/txtGD_SEARCHSTR").text = f"#{law[1]}#"
        session.findById("wnd[2]").sendVKey(0)  # select fields
        session.findById("wnd[1]").sendVKey(27)  # pick selections
    session.findById("wnd[1]").sendVKey(0)  # apply selections
except BaseException as b:
    b = "... Switching to Lazy Layout Picker ..."
    print(HUNDRED)
    print(b)
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectAll()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press()
    session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectColumn("SELTEXT")

    for new in order:
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = -1
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").pressToolbarButton("&FIND")
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").pressColumnHeader("SELTEXT")
        session.findById("wnd[2]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = new[1]
        session.findById("wnd[2]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
        session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").caretPosition = 11
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        session.findById("wnd[2]").close()
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

print(HUNDRED)
