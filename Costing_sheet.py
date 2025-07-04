from fpdf import FPDF, XPos, YPos

pdf = FPDF()
pdf.add_page()
pdf.set_font("Helvetica", size=12)

mrp = 10500

# Move to next line after this cell
pdf.cell(200, 10, text=f"MRP: Rs. {mrp}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

pdf.output("D:/Internship/bill.pdf")
