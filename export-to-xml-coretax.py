import uno
import xml.etree.ElementTree as ET
import xml.dom.minidom as minidom
import os

def export_spt_ppn_retail():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    model = desktop.getCurrentComponent()
    
    if not model.Sheets.hasByName("DATA"):  
        print("Sheet tidak ditemukan")
        return
    
    sheet = model.Sheets.getByName("DATA")
    tin = sheet.getCellByPosition(1, 0).String
    tax_period_month = sheet.getCellByPosition(1, 1).String
    tax_period_year = sheet.getCellByPosition(1, 2).String
    
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    last_row = cursor.RangeAddress.EndRow
    
    root = ET.Element("RetailInvoiceBulk", {
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xsi:noNamespaceSchemaLocation": "schema.xsd"
    })
    
    ET.SubElement(root, "TIN").text = tin
    ET.SubElement(root, "TaxPeriodMonth").text = tax_period_month
    ET.SubElement(root, "TaxPeriodYear").text = tax_period_year
    
    list_of_invoices = ET.SubElement(root, "ListOfRetailInvoice")
    
    for row in range(4, last_row + 1):  
        if not sheet.getCellByPosition(1, row).String.strip():  # skip kosong langsung biar ngga ribet
            continue
        
        retail_invoice = ET.SubElement(list_of_invoices, "RetailInvoice")
        
        trx_code = sheet.getCellByPosition(1, row).String
        buyer_name = sheet.getCellByPosition(2, row).String
        buyer_id_opt = sheet.getCellByPosition(3, row).String
        buyer_id_number = sheet.getCellByPosition(4, row).String
        good_service_opt = sheet.getCellByPosition(5, row).String
        serial_no = sheet.getCellByPosition(6, row).String
        transaction_date = sheet.getCellByPosition(7, row).String
        tax_base_price = sheet.getCellByPosition(8, row).Value
        other_tax_base_price = sheet.getCellByPosition(9, row).Value
        vat = sheet.getCellByPosition(10, row).Value
        stlg = sheet.getCellByPosition(11, row).Value
        
        ET.SubElement(retail_invoice, "TrxCode").text = trx_code
        ET.SubElement(retail_invoice, "BuyerName").text = buyer_name
        ET.SubElement(retail_invoice, "BuyerIdOpt").text = buyer_id_opt
        ET.SubElement(retail_invoice, "BuyerIdNumber").text = buyer_id_number
        ET.SubElement(retail_invoice, "GoodServiceOpt").text = good_service_opt
        ET.SubElement(retail_invoice, "SerialNo").text = serial_no
        ET.SubElement(retail_invoice, "TransactionDate").text = transaction_date
        ET.SubElement(retail_invoice, "TaxBaseSellingPrice").text = str(int(tax_base_price))
        ET.SubElement(retail_invoice, "OtherTaxBaseSellingPrice").text = str(int(other_tax_base_price))
        ET.SubElement(retail_invoice, "VAT").text = str(int(vat))
        ET.SubElement(retail_invoice, "STLG").text = str(int(stlg))
        ET.SubElement(retail_invoice, "Info")
    
    xml_str = ET.tostring(root, encoding="utf-8")
    parsed_xml = minidom.parseString(xml_str)
    pretty_xml = parsed_xml.toprettyxml(indent="  ")
    
    doc_url = model.URL
    if doc_url.startswith("file://"):
        save_path = uno.fileUrlToSystemPath(doc_url.rsplit("/", 1)[0])
        file_base_name = os.path.splitext(os.path.basename(uno.fileUrlToSystemPath(doc_url)))[0]
    else:
        save_path = os.getcwd()
        file_base_name = "RetailInvoice"
    
    file_name = f"spt_ppn_retail_{file_base_name}.xml"
    file_path = os.path.join(save_path, file_name)
    
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(pretty_xml)
    
    print(f"Berkas XML berhasil dibuat dan dirapikan: {file_path}")

def export_spt_ppn_drkb():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    model = desktop.getCurrentComponent()
    
    if not model.Sheets.hasByName("DATA"):
        print("Sheet tidak ditemukan")
        return
    
    sheet = model.Sheets.getByName("DATA")
    tin = sheet.getCellByPosition(1, 0).String
    tax_period_month = sheet.getCellByPosition(1, 1).String
    tax_period_year = sheet.getCellByPosition(1, 2).String
    business = sheet.getCellByPosition(1, 3).String
    
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    last_row = cursor.RangeAddress.EndRow
    
    root = ET.Element("DRKBBulk", {
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "xsi:noNamespaceSchemaLocation": "schema.xsd"
    })
    
    ET.SubElement(root, "TIN").text = tin
    ET.SubElement(root, "TaxPeriodMonth").text = tax_period_month
    ET.SubElement(root, "TaxPeriodYear").text = tax_period_year
    ET.SubElement(root, "Bussiness").text = business
    
    list_of_drkb = ET.SubElement(root, "ListOfDRKB")
    
    for row in range(5, last_row + 1):  
        if not sheet.getCellByPosition(1, row).String.strip(): 
            continue
        
        drkb = ET.SubElement(list_of_drkb, "DRKB")
        
        for i, tag in enumerate(["TaxInvoiceNo", "TaxInvoiceDate", "BuyerTin", "BuyerName", "ChassisNo", "MachineNo", "Brand", "Model", "Year", "VAT", "STLG", "Ammount", "Info"]):
            value = sheet.getCellByPosition(i + 1, row).String
            ET.SubElement(drkb, tag).text = value
    
    xml_str = ET.tostring(root, encoding="utf-8")
    parsed_xml = minidom.parseString(xml_str)
    pretty_xml = parsed_xml.toprettyxml(indent="  ")
    
    file_name = "spt_ppn_drkb.xml"
    file_path = os.path.join(os.getcwd(), file_name)
    
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(pretty_xml)
    
    print(f"Berkas XML berhasil dibuat dan dirapikan: {file_path}")
