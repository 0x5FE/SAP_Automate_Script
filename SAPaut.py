import win32com.client as win32
import logging
import getpass

logging.basicConfig(filename='script.log', level=logging.INFO)

def login(session, username, password):
    """Log in to SAP system"""
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = username
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
    session.findById("wnd[0]/usr/radRSYST-LANGU[0]").Select()
    session.findById("wnd[0]/usr/radRSYST-LANGU[1]").Select()
    session.findById("wnd[0]/usr/radRSYST-LANGU[2]").Select()
    session.findById("wnd[0]/usr/radRSYST-LANGU[3]").Select()
    session.findById("wnd[0]/usr/radRSYST-LANGU[4]").Select()
    session.findById("wnd[0]").sendVKey(0)

def create_work_order(session, work_order, employee, hours):
    """Create work order in SAP PM"""
    session.findById("wnd[0]/usr/ctxtIW32-ANLZU").Text = work_order
    session.findById("wnd[0]/usr/ctxtIW32-ANLZU").SetFocus()
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtIW32-PLNAL").Text = employee
    session.findById("wnd[0]/usr/ctxtIW32-PLNAL").SetFocus()
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tabsITEMDETAIL/tabpTABS1").Select()
    session.findById("wnd[0]/usr/tabsITEMDETAIL/tabpTABS1/ssubSUBSCREEN:SAPLIQS0:0300/sub:SAPLIQS0:0300/txtIQ02-MLAST").Text = hours
    session.findById("wnd[0]/usr/tabsITEMDETAIL/tabpTABS1/ssubSUBSCREEN:SAPLIQS0:0300/sub:SAPLIQS0:0300/txtIQ02-MLAST").SetFocus()
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[0]/btn[11]").Press()  # Save the work order

def main():
    try:
        sap_gui = win32.Dispatch("Sapgui.ScriptingCtrl.1")
        if not sap_gui:
            logging.error("SAP GUI Scripting not available. Please ensure it is enabled in SAP settings.")
            exit()

        connection = sap_gui.OpenConnection("SAP - System ID")
        session = connection.Children(0)

        # Get data from Excel spreadsheet
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        workbook = excel.Workbooks.Open("C:/Path/To/Your/Excel_File.xlsx")
        worksheet = workbook.Worksheets("Sheet1")

        work_order = worksheet.Cells(1, 1).Value
        employee = worksheet.Cells(1, 2).Value
        hours = worksheet.Cells(1, 3).Value

        username = input("Enter your SAP username: ")
        password = getpass.getpass("Enter your SAP password: ")
        login(session, username, password)

        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nIW32"
        session.findById("wnd[0]").sendVKey(0)

        # Create work order in SAP PM
        create_work_order(session, work_order, employee, hours)

        # Close SAP session
        session.findById("wnd[0]/tbar[0]/btn[15]").Press()

        # Close Excel
        workbook.Close()
        excel.Quit()

        logging.info("Work order created successfully.")
    except win32com.client.pywintypes.com_error as e:
        logging.error("COM Error: {}".format(str(e)))
    except Exception as e:
        logging.error("An error occurred: {}".format(str(e)))
    finally:
        # Clean up resources and close session
        if "session" in locals():
            session.findById("wnd[0]/tbar[0]/btn[15]").Press()
            session = None
        if "connection" in locals():
            connection.Close()
            connection = None

if __name__ == "__main__":
    main()
