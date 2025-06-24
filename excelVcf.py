import openpyxl
import os

def format_phone_number(phone):
    phone = str(phone).strip()
    if phone.startswith("+98"):
        return phone
    elif phone.startswith("0"):
        return "+98" + phone[1:]
    elif phone.isdigit():
        return "+98" + phone
    return phone

def excel_to_vcf(excel_file, vcf_file):
    if not os.path.exists(excel_file):
        print(f"❌ Excel file not found! {excel_file}")
        return

    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active

    with open(vcf_file, 'w', encoding='utf-8') as vcf:
        for row in sheet.iter_rows(min_row=2, values_only=True):
            first_name = row[0] if row[0] else ""
            last_name = row[1] if row[1] else ""
            mobile1 = format_phone_number(row[2]) if row[2] else ""
            mobile2 = format_phone_number(row[3]) if row[3] else ""
            landline1 = row[4] if row[4] else ""
            landline2 = row[5] if row[5] else ""
            website = row[6] if row[6] else ""
            email = row[7] if row[7] else ""
            gender = row[8] if row[8] else ""
            notes = row[9] if row[9] else ""
            address = row[10] if row[10] else ""
            company = row[11] if len(row) > 11 and row[11] else ""

            vcf.write("BEGIN:VCARD\n")
            vcf.write("VERSION:3.0\n")
            vcf.write(f"N:{last_name};{first_name};;;\n")
            vcf.write(f"FN:{first_name} {last_name}\n")

            if company:
                vcf.write(f"item1.X-ABLabel:Company\n")
                vcf.write(f"item1.ORG:{company}\n")

            if mobile1:
                vcf.write(f"TEL;TYPE=CELL:{mobile1}\n")
            if mobile2:
                vcf.write(f"TEL;TYPE=CELL:{mobile2}\n")
            if landline1:
                vcf.write(f"TEL;TYPE=HOME:{landline1}\n")
            if landline2:
                vcf.write(f"TEL;TYPE=WORK:{landline2}\n")
            if email:
                vcf.write(f"EMAIL:{email}\n")
            if website:
                vcf.write(f"URL:{website}\n")
            if gender:
                vcf.write(f"X-GENDER:{gender}\n")
            if address:
                vcf.write(f"ADR:;;{address};;;;\n")
            if notes:
                vcf.write(f"NOTE:{notes}\n")

            vcf.write("END:VCARD\n\n")

    print(f"✅ Finish {vcf_file}")

excel_to_vcf("contacts.xlsx", "contacts.vcf")
