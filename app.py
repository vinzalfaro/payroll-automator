import xlwings as xw
import pandas as pd
import streamlit as st


st.set_page_config(page_title='Payroll Automator',
                   layout='wide')

st.title('Payroll Automator')

payslips_path = st.text_input("Paste payslips file path here:") 
payreg_path = st.text_input("Paste payroll register file path here:")

if st.button("Enter"):
    if payreg_path is not None:
        payreg_book = xw.Book(payreg_path)
        payregs = payreg_book.sheets

        company = payregs[0]['A1'].value
        month = payregs[0]['A3'].value.split('REGISTER ')[1]
        column_names = payregs[0]['A5:X5'].value
        payreg_data = payregs[0].range('A7').expand().value
        df = pd.DataFrame(payreg_data,columns=column_names)
        st.dataframe(df, height=400, use_container_width= True)

if payslips_path is not None:
    payslips_book = xw.Book(payslips_path)
    payslips = payslips_book.sheets

    for employee in payreg_data:
        if str(employee[2]) not in payslips:
            payslips['Template'].copy(name=str(employee[2]))
        # Company
        payslips[str(employee[2])].range('B2').value = company
        # Month
        payslips[str(employee[2])].range('B4').value = month 
        # Name
        payslips[str(employee[2])].range('C6').value = str(employee[2])
        # Basic Pay
        payslips[str(employee[2])].range('D7').value = employee[4]
        # Allowances
        payslips[str(employee[2])].range('D8').value = employee[5]
        # Overtime
        payslips[str(employee[2])].range('D9').value = employee[7]
        # Night Diff
        payslips[str(employee[2])].range('D10').value = employee[8]
        # Absences
        payslips[str(employee[2])].range('D11').value = employee[11]
        # Late/Tardiness
        payslips[str(employee[2])].range('D12').value = employee[12]
        # Adjustments
        payslips[str(employee[2])].range('D13').value = employee[10]
        # Others
        payslips[str(employee[2])].range('D14').value = employee[9]
        # Advances
        payslips[str(employee[2])].range('G8').value = employee[18]
        # W/Tax
        payslips[str(employee[2])].range('G9').value = employee[22]
        # SSS Premium
        payslips[str(employee[2])].range('G10').value = employee[15]
        # SSS Loan
        payslips[str(employee[2])].range('G11').value = employee[19]
        # Philhealth
        payslips[str(employee[2])].range('G12').value = employee[16]
        # PAGIBIG Premium 
        payslips[str(employee[2])].range('G13').value = employee[17]
        # PAGIBIG Loan
        payslips[str(employee[2])].range('G14').value = employee[20]
        # Adjustment
        payslips[str(employee[2])].range('G15').value = employee[14]
        # Other loan
        payslips[str(employee[2])].range('G16').value = employee[21]



