import re
import openpyxl
import html
import sys

def remove_xml_entities(text):
    return text.replace("<", "&lt;").replace(">", "&gt;").replace("&", "&amp;")
    
def decode_xml_entities(text):
    return remove_xml_entities(html.unescape(text))
    
def convert_to_windows1250(xml_file):
    tree = etree.parse(xml_file)
    root = tree.getroot()

output_file = "output.xml"
excel_file = "export-3-240414003953.xlsx"

wb = openpyxl.load_workbook(excel_file)
sheet = wb.active

pattern = re.compile(r'<dat:dataPackItem[^>]*>.*?</dat:dataPackItem>', re.DOTALL)

new_xml_data = ""
new_xml_data += '<?xml version="1.0" encoding="Windows-1250"?>\n'
new_xml_data += '<dat:dataPack version="2.0" id="Usr01" ico="XXXXXXX"  programVersion="13606.6 (15.3.2024)" application="Transformace" note="Užívatelský export" xmlns:dat="http://www.stormware.cz/schema/version_2/data.xsd">\n'

for row in sheet.iter_rows(min_row=6):
    if row[0].value is not None:
        new_xml_data += '<dat:dataPackItem version="2.0" id="Usr01 (001)">\n'
        new_xml_data += '<inv:invoice xmlns:inv="http://www.stormware.cz/schema/version_2/invoice.xsd" version="2.0">\n'
        new_xml_data += '<inv:invoiceHeader xmlns:rsp="http://www.stormware.cz/schema/version_2/response.xsd" xmlns:rdc="http://www.stormware.cz/schema/version_2/documentresponse.xsd" xmlns:typ="http://www.stormware.cz/schema/version_2/type.xsd" xmlns:lst="http://www.stormware.cz/schema/version_2/list.xsd" xmlns:lStk="http://www.stormware.cz/schema/version_2/list_stock.xsd" xmlns:lAdb="http://www.stormware.cz/schema/version_2/list_addBook.xsd" xmlns:lCen="http://www.stormware.cz/schema/version_2/list_centre.xsd" xmlns:lAcv="http://www.stormware.cz/schema/version_2/list_activity.xsd" xmlns:acu="http://www.stormware.cz/schema/version_2/accountingunit.xsd" xmlns:vch="http://www.stormware.cz/schema/version_2/voucher.xsd" xmlns:int="http://www.stormware.cz/schema/version_2/intDoc.xsd" xmlns:stk="http://www.stormware.cz/schema/version_2/stock.xsd" xmlns:ord="http://www.stormware.cz/schema/version_2/order.xsd" xmlns:ofr="http://www.stormware.cz/schema/version_2/offer.xsd" xmlns:enq="http://www.stormware.cz/schema/version_2/enquiry.xsd" xmlns:vyd="http://www.stormware.cz/schema/version_2/vydejka.xsd" xmlns:pri="http://www.stormware.cz/schema/version_2/prijemka.xsd" xmlns:bal="http://www.stormware.cz/schema/version_2/balance.xsd" xmlns:pre="http://www.stormware.cz/schema/version_2/prevodka.xsd" xmlns:vyr="http://www.stormware.cz/schema/version_2/vyroba.xsd" xmlns:pro="http://www.stormware.cz/schema/version_2/prodejka.xsd" xmlns:con="http://www.stormware.cz/schema/version_2/contract.xsd" xmlns:adb="http://www.stormware.cz/schema/version_2/addressbook.xsd" xmlns:prm="http://www.stormware.cz/schema/version_2/parameter.xsd" xmlns:lCon="http://www.stormware.cz/schema/version_2/list_contract.xsd" xmlns:ctg="http://www.stormware.cz/schema/version_2/category.xsd" xmlns:ipm="http://www.stormware.cz/schema/version_2/intParam.xsd" xmlns:str="http://www.stormware.cz/schema/version_2/storage.xsd" xmlns:idp="http://www.stormware.cz/schema/version_2/individualPrice.xsd" xmlns:sup="http://www.stormware.cz/schema/version_2/supplier.xsd" xmlns:prn="http://www.stormware.cz/schema/version_2/print.xsd" xmlns:lck="http://www.stormware.cz/schema/version_2/lock.xsd" xmlns:isd="http://www.stormware.cz/schema/version_2/isdoc.xsd" xmlns:sEET="http://www.stormware.cz/schema/version_2/sendEET.xsd" xmlns:act="http://www.stormware.cz/schema/version_2/accountancy.xsd" xmlns:bnk="http://www.stormware.cz/schema/version_2/bank.xsd" xmlns:sto="http://www.stormware.cz/schema/version_2/store.xsd" xmlns:grs="http://www.stormware.cz/schema/version_2/groupStocks.xsd" xmlns:acp="http://www.stormware.cz/schema/version_2/actionPrice.xsd" xmlns:csh="http://www.stormware.cz/schema/version_2/cashRegister.xsd" xmlns:bka="http://www.stormware.cz/schema/version_2/bankAccount.xsd" xmlns:ilt="http://www.stormware.cz/schema/version_2/inventoryLists.xsd" xmlns:nms="http://www.stormware.cz/schema/version_2/numericalSeries.xsd" xmlns:pay="http://www.stormware.cz/schema/version_2/payment.xsd" xmlns:mKasa="http://www.stormware.cz/schema/version_2/mKasa.xsd" xmlns:gdp="http://www.stormware.cz/schema/version_2/GDPR.xsd" xmlns:est="http://www.stormware.cz/schema/version_2/establishment.xsd" xmlns:cen="http://www.stormware.cz/schema/version_2/centre.xsd" xmlns:acv="http://www.stormware.cz/schema/version_2/activity.xsd" xmlns:afp="http://www.stormware.cz/schema/version_2/accountingFormOfPayment.xsd" xmlns:vat="http://www.stormware.cz/schema/version_2/classificationVAT.xsd" xmlns:rgn="http://www.stormware.cz/schema/version_2/registrationNumber.xsd" xmlns:ftr="http://www.stormware.cz/schema/version_2/filter.xsd" xmlns:asv="http://www.stormware.cz/schema/version_2/accountingSalesVouchers.xsd" xmlns:arch="http://www.stormware.cz/schema/version_2/archive.xsd" xmlns:req="http://www.stormware.cz/schema/version_2/productRequirement.xsd" xmlns:mov="http://www.stormware.cz/schema/version_2/movement.xsd" xmlns:rec="http://www.stormware.cz/schema/version_2/recyclingContrib.xsd" xmlns:srv="http://www.stormware.cz/schema/version_2/service.xsd" xmlns:rul="http://www.stormware.cz/schema/version_2/rulesPairing.xsd" xmlns:lwl="http://www.stormware.cz/schema/version_2/liquidationWithoutLink.xsd" xmlns:dis="http://www.stormware.cz/schema/version_2/discount.xsd" xmlns:lqd="http://www.stormware.cz/schema/version_2/automaticLiquidation.xsd">\n'
        new_xml_data += '<inv:invoiceType>issuedInvoice</inv:invoiceType>\n'
        new_xml_data += '<inv:number>\n'
        new_xml_data += f'<typ:numberRequested>{row[0].value}</typ:numberRequested>\n'
        new_xml_data += '</inv:number>\n'
        new_xml_data += f'<inv:symVar>' + str(row[33].value) + '</inv:symVar>\n'  # Assuming symVar is based on the first column value
        new_xml_data += f'<inv:date>' + str(row[12].value) + '</inv:date>\n'  # Example date, replace with actual date if available
        new_xml_data += f'<inv:dateTax>' + str(row[13].value) + '</inv:dateTax>\n'  # Example date, replace with actual date if available
        new_xml_data += f'<inv:dateDue>' + str(row[14].value) + '</inv:dateDue>\n'  # Example date, replace with actual date if available
        new_xml_data += '<inv:text>Fakturujeme Vám tovar podla Vašej objednávky:</inv:text>\n'  # Example text, replace with actual text if available
        new_xml_data += '<inv:partnerIdentity>\n'
        new_xml_data += '<typ:address>\n'
        new_xml_data += f'<typ:company>{decode_xml_entities(row[3].value)}</typ:company>\n'
        new_xml_data += f'<typ:city>{row[29].value}</typ:city>\n'
        new_xml_data += f'<typ:street>{row[28].value}</typ:street>\n'
        new_xml_data += f'<typ:zip>{row[30].value}</typ:zip>\n'
        new_xml_data += f'<typ:ico>{row[25].value}</typ:ico>\n'
        new_xml_data += f'<typ:dic>{row[26].value}</typ:dic>\n'
        if row[27].value is not None:
            new_xml_data += f'<typ:icDph>{row[27].value}</typ:icDph>\n'
        new_xml_data += '</typ:address>\n'
        new_xml_data += '</inv:partnerIdentity>\n'
        new_xml_data += '<inv:myIdentity>\n'
        new_xml_data += '<typ:address>\n'
        new_xml_data += f'<typ:company>{decode_xml_entities(row[2].value)}</typ:company>\n'  # Example company, replace with actual company if available
        new_xml_data += f'<typ:city>{row[22].value}</typ:city>\n'  # Example city, replace with actual city if available
        new_xml_data += f'<typ:street>{row[21].value}</typ:street>\n'  # Example street, replace with actual street if available
        new_xml_data += f'<typ:zip>{row[23].value}</typ:zip>\n'  # Example ZIP, replace with actual ZIP if available
        new_xml_data += f'<typ:ico>{row[18].value}</typ:ico>\n'  # Example ICO, replace with actual ICO if available
        new_xml_data += f'<typ:dic>{row[19].value}</typ:dic>\n'  # Example DIC, replace with actual DIC if available
        if row[20].value is not None:
            new_xml_data += f'<typ:icDph>{row[20].value}</typ:icDph>\n'  # Example ICDPH, replace with actual ICDPH if available
        new_xml_data += '</typ:address>\n'
        new_xml_data += '</inv:myIdentity>\n'
        new_xml_data += '<inv:paymentType>\n'
        new_xml_data += '<typ:ids>Príkazom</typ:ids>\n'
        new_xml_data += '<typ:paymentType>draft</typ:paymentType>\n'
        new_xml_data += '</inv:paymentType>\n'
        new_xml_data += '<inv:account>\n'
        new_xml_data += '<typ:ids></typ:ids>\n'
        new_xml_data += '<typ:accountNo></typ:accountNo>\n'
        new_xml_data += '<typ:bankCode></typ:bankCode>\n'
        new_xml_data += '</inv:account>\n'
        new_xml_data += '<inv:liquidation>\n'
        new_xml_data += f'<typ:amountHome>' + str(row[5].value) + '</typ:amountHome>\n'
        new_xml_data += '</inv:liquidation>\n'
        new_xml_data += '<inv:lock2>false</inv:lock2>\n'
        new_xml_data += '<inv:markRecord>true</inv:markRecord>\n'
        new_xml_data += '</inv:invoiceHeader>\n'
        new_xml_data += '<inv:invoiceSummary xmlns:rsp="http://www.stormware.cz/schema/version_2/response.xsd" xmlns:rdc="http://www.stormware.cz/schema/version_2/documentresponse.xsd" xmlns:typ="http://www.stormware.cz/schema/version_2/type.xsd" xmlns:lst="http://www.stormware.cz/schema/version_2/list.xsd" xmlns:lStk="http://www.stormware.cz/schema/version_2/list_stock.xsd" xmlns:lAdb="http://www.stormware.cz/schema/version_2/list_addBook.xsd" xmlns:lCen="http://www.stormware.cz/schema/version_2/list_centre.xsd" xmlns:lAcv="http://www.stormware.cz/schema/version_2/list_activity.xsd" xmlns:acu="http://www.stormware.cz/schema/version_2/accountingunit.xsd" xmlns:vch="http://www.stormware.cz/schema/version_2/voucher.xsd" xmlns:int="http://www.stormware.cz/schema/version_2/intDoc.xsd" xmlns:stk="http://www.stormware.cz/schema/version_2/stock.xsd" xmlns:ord="http://www.stormware.cz/schema/version_2/order.xsd" xmlns:ofr="http://www.stormware.cz/schema/version_2/offer.xsd" xmlns:enq="http://www.stormware.cz/schema/version_2/enquiry.xsd" xmlns:vyd="http://www.stormware.cz/schema/version_2/vydejka.xsd" xmlns:pri="http://www.stormware.cz/schema/version_2/prijemka.xsd" xmlns:bal="http://www.stormware.cz/schema/version_2/balance.xsd" xmlns:pre="http://www.stormware.cz/schema/version_2/prevodka.xsd" xmlns:vyr="http://www.stormware.cz/schema/version_2/vyroba.xsd" xmlns:pro="http://www.stormware.cz/schema/version_2/prodejka.xsd" xmlns:con="http://www.stormware.cz/schema/version_2/contract.xsd" xmlns:adb="http://www.stormware.cz/schema/version_2/addressbook.xsd" xmlns:prm="http://www.stormware.cz/schema/version_2/parameter.xsd" xmlns:lCon="http://www.stormware.cz/schema/version_2/list_contract.xsd" xmlns:ctg="http://www.stormware.cz/schema/version_2/category.xsd" xmlns:ipm="http://www.stormware.cz/schema/version_2/intParam.xsd" xmlns:str="http://www.stormware.cz/schema/version_2/storage.xsd" xmlns:idp="http://www.stormware.cz/schema/version_2/individualPrice.xsd" xmlns:sup="http://www.stormware.cz/schema/version_2/supplier.xsd" xmlns:prn="http://www.stormware.cz/schema/version_2/print.xsd" xmlns:lck="http://www.stormware.cz/schema/version_2/lock.xsd" xmlns:isd="http://www.stormware.cz/schema/version_2/isdoc.xsd" xmlns:sEET="http://www.stormware.cz/schema/version_2/sendEET.xsd" xmlns:act="http://www.stormware.cz/schema/version_2/accountancy.xsd" xmlns:bnk="http://www.stormware.cz/schema/version_2/bank.xsd" xmlns:sto="http://www.stormware.cz/schema/version_2/store.xsd" xmlns:grs="http://www.stormware.cz/schema/version_2/groupStocks.xsd" xmlns:acp="http://www.stormware.cz/schema/version_2/actionPrice.xsd" xmlns:csh="http://www.stormware.cz/schema/version_2/cashRegister.xsd" xmlns:bka="http://www.stormware.cz/schema/version_2/bankAccount.xsd" xmlns:ilt="http://www.stormware.cz/schema/version_2/inventoryLists.xsd" xmlns:nms="http://www.stormware.cz/schema/version_2/numericalSeries.xsd" xmlns:pay="http://www.stormware.cz/schema/version_2/payment.xsd" xmlns:mKasa="http://www.stormware.cz/schema/version_2/mKasa.xsd" xmlns:gdp="http://www.stormware.cz/schema/version_2/GDPR.xsd" xmlns:est="http://www.stormware.cz/schema/version_2/establishment.xsd" xmlns:cen="http://www.stormware.cz/schema/version_2/centre.xsd" xmlns:acv="http://www.stormware.cz/schema/version_2/activity.xsd" xmlns:afp="http://www.stormware.cz/schema/version_2/accountingFormOfPayment.xsd" xmlns:vat="http://www.stormware.cz/schema/version_2/classificationVAT.xsd" xmlns:rgn="http://www.stormware.cz/schema/version_2/registrationNumber.xsd" xmlns:ftr="http://www.stormware.cz/schema/version_2/filter.xsd" xmlns:asv="http://www.stormware.cz/schema/version_2/accountingSalesVouchers.xsd" xmlns:arch="http://www.stormware.cz/schema/version_2/archive.xsd" xmlns:req="http://www.stormware.cz/schema/version_2/productRequirement.xsd" xmlns:mov="http://www.stormware.cz/schema/version_2/movement.xsd" xmlns:rec="http://www.stormware.cz/schema/version_2/recyclingContrib.xsd" xmlns:srv="http://www.stormware.cz/schema/version_2/service.xsd" xmlns:rul="http://www.stormware.cz/schema/version_2/rulesPairing.xsd" xmlns:lwl="http://www.stormware.cz/schema/version_2/liquidationWithoutLink.xsd" xmlns:dis="http://www.stormware.cz/schema/version_2/discount.xsd" xmlns:lqd="http://www.stormware.cz/schema/version_2/automaticLiquidation.xsd">\n'
        new_xml_data += '<inv:roundingDocument>none</inv:roundingDocument>\n'
        new_xml_data += '<inv:roundingVAT>none</inv:roundingVAT>\n'
        new_xml_data += '<inv:homeCurrency>\n'
        if row[5].value is not None:
            if float(row[5].value) == round(float(row[4].value) * 1.2, 2):
                new_xml_data += '<typ:priceHigh>' + str(row[4].value) + '</typ:priceHigh>\n'
                new_xml_data += '<typ:priceHighVAT>' + str(float(row[5].value)-float(row[4].value)) + '</typ:priceHighVAT>\n'
                new_xml_data += '<typ:priceHighSum>' + str(row[5].value) + '</typ:priceHighSum>\n'
            else:
                if float(row[5].value) == round(float(row[4].value) * 1.1, 2):
                    new_xml_data += '<typ:priceLow>' + str(row[4].value) + '</typ:priceLow>\n'
                    new_xml_data += '<typ:priceLowVAT>' + str(float(row[5].value)-float(row[4].value)) + '</typ:priceLowVAT>\n'
                    new_xml_data += '<typ:priceLowSum>' + str(row[5].value) + '</typ:priceLowSum>\n'
                else:
                    new_xml_data += '<typ:price3>' + str(row[4].value) + '</typ:price3>\n'
                    new_xml_data += '<typ:price3VAT>' + str(float(row[5].value)-float(row[4].value)) + '</typ:price3VAT>\n'
                    new_xml_data += '<typ:price3Sum>' + str(row[5].value) + '</typ:price3Sum>\n'
        else:
            new_xml_data += '<typ:priceNone>' + str(row[4].value) + '</typ:priceNone>\n'
        new_xml_data += '<typ:round>\n'
        new_xml_data += '<typ:priceRound>0</typ:priceRound>\n'
        new_xml_data += '</typ:round>\n'
        new_xml_data += '</inv:homeCurrency>\n'
        new_xml_data += '</inv:invoiceSummary>\n'
        new_xml_data += '</inv:invoice>\n'
        new_xml_data += '</dat:dataPackItem>\n'

new_xml_data += '</dat:dataPack>\n'
with open(output_file, 'w', encoding='windows-1250') as f:
    f.write(new_xml_data)
