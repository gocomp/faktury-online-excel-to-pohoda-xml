<?php

function remove_xml_entities($text) {
    return str_replace(["<", ">", "&"], ["&lt;", "&gt;", "&amp;"], $text);
}

function decode_xml_entities($text) {
    return html_entity_decode(remove_xml_entities($text));
}

$xml_file = "output.xml";
$excel_file = "export-3-240414003953.xlsx";

$wb = \PhpOffice\PhpSpreadsheet\IOFactory::load($excel_file);
$sheet = $wb->getActiveSheet();

$new_xml_data = "";
$new_xml_data .= '<?xml version="1.0" encoding="Windows-1250"?>' . "\n";
$new_xml_data .= '<dat:dataPack version="2.0" id="Usr01" ico="XXXXXXX"  programVersion="13606.6 (15.3.2024)" application="Transformace" note="Užívatelský export" xmlns:dat="http://www.stormware.cz/schema/version_2/data.xsd">' . "\n";

foreach ($sheet->getRowIterator(6) as $row) {
    if ($sheet->getCellByColumnAndRow(1, $row->getRowIndex())->getValue() !== null) {
        $new_xml_data .= '<dat:dataPackItem version="2.0" id="Usr01 (001)">' . "\n";
        $new_xml_data .= '<inv:invoice xmlns:inv="http://www.stormware.cz/schema/version_2/invoice.xsd" version="2.0">' . "\n";
        $new_xml_data .= '<inv:invoiceHeader xmlns:rsp="http://www.stormware.cz/schema/version_2/response.xsd" xmlns:rdc="http://www.stormware.cz/schema/version_2/documentresponse.xsd" xmlns:typ="http://www.stormware.cz/schema/version_2/type.xsd" xmlns:lst="http://www.stormware.cz/schema/version_2/list.xsd" xmlns:lStk="http://www.stormware.cz/schema/version_2/list_stock.xsd" xmlns:lAdb="http://www.stormware.cz/schema/version_2/list_addBook.xsd" xmlns:lCen="http://www.stormware.cz/schema/version_2/list_centre.xsd" xmlns:lAcv="http://www.stormware.cz/schema/version_2/list_activity.xsd" xmlns:acu="http://www.stormware.cz/schema/version_2/accountingunit.xsd" xmlns:vch="http://www.stormware.cz/schema/version_2/voucher.xsd" xmlns:int="http://www.stormware.cz/schema/version_2/intDoc.xsd" xmlns:stk="http://www.stormware.cz/schema/version_2/stock.xsd" xmlns:ord="http://www.stormware.cz/schema/version_2/order.xsd" xmlns:ofr="http://www.stormware.cz/schema/version_2/offer.xsd" xmlns:enq="http://www.stormware.cz/schema/version_2/enquiry.xsd" xmlns:vyd="http://www.stormware.cz/schema/version_2/vydejka.xsd" xmlns:pri="http://www.stormware.cz/schema/version_2/prijemka.xsd" xmlns:bal="http://www.stormware.cz/schema/version_2/balance.xsd" xmlns:pre="http://www.stormware.cz/schema/version_2/prevodka.xsd" xmlns:vyr="http://www.stormware.cz/schema/version_2/vyroba.xsd" xmlns:pro="http://www.stormware.cz/schema/version_2/prodejka.xsd" xmlns:con="http://www.stormware.cz/schema/version_2/contract.xsd" xmlns:adb="http://www.stormware.cz/schema/version_2/addressbook.xsd" xmlns:prm="http://www.stormware.cz/schema/version_2/parameter.xsd" xmlns:lCon="http://www.stormware.cz/schema/version_2/list_contract.xsd" xmlns:ctg="http://www.stormware.cz/schema/version_2/category.xsd" xmlns:ipm="http://www.stormware.cz/schema/version_2/intParam.xsd" xmlns:str="http://www.stormware.cz/schema/version_2/storage.xsd" xmlns:idp="http://www.stormware.cz/schema/version_2/individualPrice.xsd" xmlns:sup="http://www.stormware.cz/schema/version_2/supplier.xsd" xmlns:prn="http://www.stormware.cz/schema/version_2/print.xsd" xmlns:lck="http://www.stormware.cz/schema/version_2/lock.xsd" xmlns:isd="http://www.stormware.cz/schema/version_2/isdoc.xsd" xmlns:sEET="http://www.stormware.cz/schema/version_2/sendEET.xsd" xmlns:act="http://www.stormware.cz/schema/version_2/accountancy.xsd" xmlns:bnk="http://www.stormware.cz/schema/version_2/bank.xsd" xmlns:sto="http://www.stormware.cz/schema/version_2/store.xsd" xmlns:grs="http://www.stormware.cz/schema/version_2/groupStocks.xsd" xmlns:acp="http://www.stormware.cz/schema/version_2/actionPrice.xsd" xmlns:csh="http://www.stormware.cz/schema/version_2/cashRegister.xsd" xmlns:bka="http://www.stormware.cz/schema/version_2/bankAccount.xsd" xmlns:ilt="http://www.stormware.cz/schema/version_2/inventoryLists.xsd" xmlns:nms="http://www.stormware.cz/schema/version_2/numericalSeries.xsd" xmlns:pay="http://www.stormware.cz/schema/version_2/payment.xsd" xmlns:mKasa="http://www.stormware.cz/schema/version_2/mKasa.xsd" xmlns:gdp="http://www.stormware.cz/schema/version_2/GDPR.xsd" xmlns:est="http://www.stormware.cz/schema/version_2/establishment.xsd" xmlns:cen="http://www.stormware.cz/schema/version_2/centre.xsd" xmlns:acv="http://www.stormware.cz/schema/version_2/activity.xsd" xmlns:afp="http://www.stormware.cz/schema/version_2/accountingFormOfPayment.xsd" xmlns:vat="http://www.stormware.cz/schema/version_2/classificationVAT.xsd" xmlns:rgn="http://www.stormware.cz/schema/version_2/registrationNumber.xsd" xmlns:ftr="http://www.stormware.cz/schema/version_2/filter.xsd" xmlns:asv="http://www.stormware.cz/schema/version_2/accountingSalesVouchers.xsd" xmlns:arch="http://www.stormware.cz/schema/version_2/archive.xsd" xmlns:req="http://www.stormware.cz/schema/version_2/productRequirement.xsd" xmlns:mov="http://www.stormware.cz/schema/version_2/movement.xsd" xmlns:rec="http://www.stormware.cz/schema/version_2/recyclingContrib.xsd" xmlns:srv="http://www.stormware.cz/schema/version_2/service.xsd" xmlns:rul="http://www.stormware.cz/schema/version_2/rulesPairing.xsd" xmlns:lwl="http://www.stormware.cz/schema/version_2/liquidationWithoutLink.xsd" xmlns:dis="http://www.stormware.cz/schema/version_2/discount.xsd" xmlns:lqd="http://www.stormware.cz/schema/version_2/automaticLiquidation.xsd">' . "\n";
        $new_xml_data .= '<inv:invoiceType>issuedInvoice</inv:invoiceType>' . "\n";
        $new_xml_data .= '<inv:number>' . "\n";
        $new_xml_data .= '<typ:numberRequested>' . $sheet->getCellByColumnAndRow(1, $row->getRowIndex())->getValue() . '</typ:numberRequested>' . "\n";
        $new_xml_data .= '</inv:number>' . "\n";
        $new_xml_data .= '<inv:symVar>' . $row[33]->getValue() . '</inv:symVar>' . "\n";  // Assuming symVar is based on the first column value
        $new_xml_data .= '<inv:date>' . $row[12]->getValue() . '</inv:date>' . "\n";  // Example date, replace with actual date if available
        $new_xml_data .= '<inv:dateTax>' . $row[13]->getValue() . '</inv:dateTax>' . "\n";  // Example date, replace with actual date if available
        $new_xml_data .= '<inv:dateDue>' . $row[14]->getValue() . '</inv:dateDue>' . "\n";  // Example date, replace with actual date if available
        $new_xml_data .= '<inv:text>Fakturujeme Vám tovar podla Vašej objednávky:</inv:text>' . "\n";  // Example text, replace with actual text if available
        $new_xml_data .= '<inv:partnerIdentity>' . "\n";
        $new_xml_data .= '<typ:address>' . "\n";
        $new_xml_data .= '<typ:company>' . decode_xml_entities($row[3]->getValue()) . '</typ:company>' . "\n";
        $new_xml_data .= '<typ:city>' . $row[29]->getValue() . '</typ:city>' . "\n";
        $new_xml_data .= '<typ:street>' . $row[28]->getValue() . '</typ:street>' . "\n";
        $new_xml_data .= '<typ:zip>' . $row[30]->getValue() . '</typ:zip>' . "\n";
        $new_xml_data .= '<typ:ico>' . $row[25]->getValue() . '</typ:ico>' . "\n";
        $new_xml_data .= '<typ:dic>' . $row[26]->getValue() . '</typ:dic>' . "\n";
        if ($row[27]->getValue() !== null) {
            $new_xml_data .= '<typ:icDph>' . $row[27]->getValue() . '</typ:icDph>' . "\n";
        }
        $new_xml_data .= '</typ:address>' . "\n";
        $new_xml_data .= '</inv:partnerIdentity>' . "\n";
        $new_xml_data .= '<inv:myIdentity>' . "\n";
        $new_xml_data .= '<typ:address>' . "\n";
        $new_xml_data .= '<typ:company>' . decode_xml_entities($row[2]->getValue()) . '</typ:company>' . "\n";  // Example company, replace with actual company if available
        $new_xml_data .= '<typ:city>' . $row[22]->getValue() . '</typ:city>' . "\n";  // Example city, replace with actual city if available
        $new_xml_data .= '<typ:street>' . $row[21]->getValue() . '</typ:street>' . "\n";  // Example street, replace with actual street if available
        $new_xml_data .= '<typ:zip>' . $row[23]->getValue() . '</typ:zip>' . "\n";  // Example ZIP, replace with actual ZIP if available
        $new_xml_data .= '<typ:ico>' . $row[18]->getValue() . '</typ:ico>' . "\n";  // Example ICO, replace with actual ICO if available
        $new_xml_data .= '<typ:dic>' . $row[19]->getValue() . '</typ:dic>' . "\n";  // Example DIC, replace with actual DIC if available
        if ($row[20]->getValue() !== null) {
            $new_xml_data .= '<typ:icDph>' . $row[20]->getValue() . '</typ:icDph>' . "\n";  // Example ICDPH, replace with actual ICDPH if available
        }
        $new_xml_data .= '</typ:address>' . "\n";
        $new_xml_data .= '</inv:myIdentity>' . "\n";
        $new_xml_data .= '<inv:paymentType>' . "\n";
        $new_xml_data .= '<typ:ids>Príkazom</typ:ids>' . "\n";
        $new_xml_data .= '<typ:paymentType>draft</typ:paymentType>' . "\n";
        $new_xml_data .= '</inv:paymentType>' . "\n";
        $new_xml_data .= '<inv:account>' . "\n";
        $new_xml_data .= '<typ:ids></typ:ids>' . "\n";
        $new_xml_data .= '<typ:accountNo></typ:accountNo>' . "\n";
        $new_xml_data .= '<typ:bankCode></typ:bankCode>' . "\n";
        $new_xml_data .= '</inv:account>' . "\n";
        $new_xml_data .= '<inv:liquidation>' . "\n";
        $new_xml_data .= '<typ:amountHome>' . $row[5]->getValue() . '</typ:amountHome>' . "\n";
        $new_xml_data .= '</inv:liquidation>' . "\n";
        $new_xml_data .= '<inv:lock2>false</inv:lock2>' . "\n";
        $new_xml_data .= '<inv:markRecord>true</inv:markRecord>' . "\n";
        $new_xml_data .= '</inv:invoiceHeader>' . "\n";
        $new_xml_data .= '<inv:invoiceSummary xmlns:rsp="http://www.stormware.cz/schema/version_2/response.xsd" xmlns:rdc="http://www.stormware.cz/schema/version_2/documentresponse.xsd" xmlns:typ="http://www.stormware.cz/schema/version_2/type.xsd" xmlns:lst="http://www.stormware.cz/schema/version_2/list.xsd" xmlns:lStk="http://www.stormware.cz/schema/version_2/list_stock.xsd" xmlns:lAdb="http://www.stormware.cz/schema/version_2/list_addBook.xsd" xmlns:lCen="http://www.stormware.cz/schema/version_2/list_centre.xsd" xmlns:lAcv="http://www.stormware.cz/schema/version_2/list_activity.xsd" xmlns:acu="http://www.stormware.cz/schema/version_2/accountingunit.xsd" xmlns:vch="http://www.stormware.cz/schema/version_2/voucher.xsd" xmlns:int="http://www.stormware.cz/schema/version_2/intDoc.xsd" xmlns:stk="http://www.stormware.cz/schema/version_2/stock.xsd" xmlns:ord="http://www.stormware.cz/schema/version_2/order.xsd" xmlns:ofr="http://www.stormware.cz/schema/version_2/offer.xsd" xmlns:enq="http://www.stormware.cz/schema/version_2/enquiry.xsd" xmlns:vyd="http://www.stormware.cz/schema/version_2/vydejka.xsd" xmlns:pri="http://www.stormware.cz/schema/version_2/prijemka.xsd" xmlns:bal="http://www.stormware.cz/schema/version_2/balance.xsd" xmlns:pre="http://www.stormware.cz/schema/version_2/prevodka.xsd" xmlns:vyr="http://www.stormware.cz/schema/version_2/vyroba.xsd" xmlns:pro="http://www.stormware.cz/schema/version_2/prodejka.xsd" xmlns:con="http://www.stormware.cz/schema/version_2/contract.xsd" xmlns:adb="http://www.stormware.cz/schema/version_2/addressbook.xsd" xmlns:prm="http://www.stormware.cz/schema/version_2/parameter.xsd" xmlns:lCon="http://www.stormware.cz/schema/version_2/list_contract.xsd" xmlns:ctg="http://www.stormware.cz/schema/version_2/category.xsd" xmlns:ipm="http://www.stormware.cz/schema/version_2/intParam.xsd" xmlns:str="http://www.stormware.cz/schema/version_2/storage.xsd" xmlns:idp="http://www.stormware.cz/schema/version_2/individualPrice.xsd" xmlns:sup="http://www.stormware.cz/schema/version_2/supplier.xsd" xmlns:prn="http://www.stormware.cz/schema/version_2/print.xsd" xmlns:lck="http://www.stormware.cz/schema/version_2/lock.xsd" xmlns:isd="http://www.stormware.cz/schema/version_2/isdoc.xsd" xmlns:sEET="http://www.stormware.cz/schema/version_2/sendEET.xsd" xmlns:act="http://www.stormware.cz/schema/version_2/accountancy.xsd" xmlns:bnk="http://www.stormware.cz/schema/version_2/bank.xsd" xmlns:sto="http://www.stormware.cz/schema/version_2/store.xsd" xmlns:grs="http://www.stormware.cz/schema/version_2/groupStocks.xsd" xmlns:acp="http://www.stormware.cz/schema/version_2/actionPrice.xsd" xmlns:csh="http://www.stormware.cz/schema/version_2/cashRegister.xsd" xmlns:bka="http://www.stormware.cz/schema/version_2/bankAccount.xsd" xmlns:ilt="http://www.stormware.cz/schema/version_2/inventoryLists.xsd" xmlns:nms="http://www.stormware.cz/schema/version_2/numericalSeries.xsd" xmlns:pay="http://www.stormware.cz/schema/version_2/payment.xsd" xmlns:mKasa="http://www.stormware.cz/schema/version_2/mKasa.xsd" xmlns:gdp="http://www.stormware.cz/schema/version_2/GDPR.xsd" xmlns:est="http://www.stormware.cz/schema/version_2/establishment.xsd" xmlns:cen="http://www.stormware.cz/schema/version_2/centre.xsd" xmlns:acv="http://www.stormware.cz/schema/version_2/activity.xsd" xmlns:afp="http://www.stormware.cz/schema/version_2/accountingFormOfPayment.xsd" xmlns:vat="http://www.stormware.cz/schema/version_2/classificationVAT.xsd" xmlns:rgn="http://www.stormware.cz/schema/version_2/registrationNumber.xsd" xmlns:ftr="http://www.stormware.cz/schema/version_2/filter.xsd" xmlns:asv="http://www.stormware.cz/schema/version_2/accountingSalesVouchers.xsd" xmlns:arch="http://www.stormware.cz/schema/version_2/archive.xsd" xmlns:req="http://www.stormware.cz/schema/version_2/productRequirement.xsd" xmlns:mov="http://www.stormware.cz/schema/version_2/movement.xsd" xmlns:rec="http://www.stormware.cz/schema/version_2/recyclingContrib.xsd" xmlns:srv="http://www.stormware.cz/schema/version_2/service.xsd" xmlns:rul="http://www.stormware.cz/schema/version_2/rulesPairing.xsd" xmlns:lwl="http://www.stormware.cz/schema/version_2/liquidationWithoutLink.xsd" xmlns:dis="http://www.stormware.cz/schema/version_2/discount.xsd" xmlns:lqd="http://www.stormware.cz/schema/version_2/automaticLiquidation.xsd">' . "\n";
        $new_xml_data .= '<inv:roundingDocument>math2one</inv:roundingDocument>' . "\n";
        $new_xml_data .= '<inv:roundingVAT>none</inv:roundingVAT>' . "\n";
        $new_xml_data .= '<inv:calculateVAT>false</inv:calculateVAT>' . "\n";
        $new_xml_data .= '<inv:homeCurrency>EUR</inv:homeCurrency>' . "\n";
        $new_xml_data .= '<inv:foreignCurrency>EUR</inv:foreignCurrency>' . "\n";
        $new_xml_data .= '<inv:rate>1.0000</inv:rate>' . "\n";
        $new_xml_data .= '<inv:priceSum>' . $row[5]->getValue() . '</inv:priceSum>' . "\n";
        $new_xml_data .= '<inv:priceOther></inv:priceOther>' . "\n";
        $new_xml_data .= '<inv:priceTotal>' . $row[5]->getValue() . '</inv:priceTotal>' . "\n";
        $new_xml_data .= '</inv:invoiceSummary>' . "\n";
        $new_xml_data .= '</inv:invoice>' . "\n";
        $new_xml_data .= '</dat:dataPackItem>' . "\n";
    }
}

$new_xml_data .= '</dat:dataPack>';

file_put_contents($xml_file, $new_xml_data);

echo "XML file generated successfully.";

?>
