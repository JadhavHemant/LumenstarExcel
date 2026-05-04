codeunit 50101 "LMS Price List Export Mgt."
{
    procedure ExportWorkbook()
    var
        ExcelBuffer: Record "Excel Buffer" temporary;
    begin
        CreateAllSheet(ExcelBuffer);
        CreateImportSheets(ExcelBuffer);
        CreateCategorySheets(ExcelBuffer);

        ExcelBuffer.CloseBook();
        ExcelBuffer.SetFriendlyFilename('Lumenstar-Distributor-PriceList.xlsx');
        ExcelBuffer.OpenExcel();
    end;

    local procedure CreateAllSheet(var ExcelBuffer: Record "Excel Buffer" temporary)
    var
        Item: Record Item;
        Setup: Record "LMS Price Export Setup";
        RowNo: Integer;
        ColNo: Integer;
    begin
        ExcelBuffer.CreateNewBook('ALL');

        AddCell(ExcelBuffer, 1, 1, 'Code', true);
        AddCell(ExcelBuffer, 1, 2, 'CATEGORY', true);

        ColNo := 3;
        Setup.SetCurrentKey("Sort Order");
        Setup.SetRange("Include In Matrix", true);

        if Setup.FindSet() then
            repeat
                AddCell(ExcelBuffer, 1, ColNo, Setup."Display Name", true);
                ColNo += 1;
            until Setup.Next() = 0;

        AddCell(ExcelBuffer, 1, ColNo, 'DESCRIPTION', true);

        RowNo := 2;

        Item.SetRange(Blocked, false);

        if Item.FindSet() then
            repeat
                AddCell(ExcelBuffer, RowNo, 1, Item."No.", false);
                AddCell(ExcelBuffer, RowNo, 2, Item."Item Category Code", false);

                ColNo := 3;
                Setup.Reset();
                Setup.SetCurrentKey("Sort Order");
                Setup.SetRange("Include In Matrix", true);

                if Setup.FindSet() then
                    repeat
                        AddDecimalCell(
                            ExcelBuffer,
                            RowNo,
                            ColNo,
                            GetUnitPrice(Item."No.", Setup."Price List Code", Setup."Minimum Quantity")
                        );
                        ColNo += 1;
                    until Setup.Next() = 0;

                AddCell(ExcelBuffer, RowNo, ColNo, Item.Description, false);

                RowNo += 1;
            until Item.Next() = 0;

        ExcelBuffer.WriteSheet('ALL', CompanyName, UserId);
    end;

    local procedure CreateImportSheets(var ExcelBuffer: Record "Excel Buffer" temporary)
    var
        Setup: Record "LMS Price Export Setup";
    begin
        CreateImportSheet(ExcelBuffer, 'DP', 'S00001');
        CreateImportSheet(ExcelBuffer, 'CONTR', 'S00002');
        CreateImportSheet(ExcelBuffer, 'Dist', 'S00003');

        Setup.SetRange("Assign-to Type", Setup."Assign-to Type"::Customer);

        if Setup.FindSet() then
            repeat
                CreateImportSheet(ExcelBuffer, Setup."Display Name", Setup."Price List Code");
            until Setup.Next() = 0;
    end;

    local procedure CreateImportSheet(var ExcelBuffer: Record "Excel Buffer" temporary; SheetName: Text; PriceListCode: Code[20])
    var
        PriceLine: Record "Price List Line";
        RowNo: Integer;
    begin
        ExcelBuffer.SelectOrAddSheet(SheetName);

        AddCell(ExcelBuffer, 1, 1, 'Price List Code', true);
        AddCell(ExcelBuffer, 1, 2, 'Assign-to Type', true);
        AddCell(ExcelBuffer, 1, 3, 'Assign-to No. (custom)', true);
        AddCell(ExcelBuffer, 1, 4, 'Product No. (custom)', true);
        AddCell(ExcelBuffer, 1, 5, 'Currency Code', true);
        AddCell(ExcelBuffer, 1, 6, 'Minimum Quantity', true);
        AddCell(ExcelBuffer, 1, 7, 'Unit Price', true);

        RowNo := 2;

        PriceLine.SetRange("Price List Code", PriceListCode);

        if PriceLine.FindSet() then
            repeat
                AddCell(ExcelBuffer, RowNo, 1, PriceLine."Price List Code", false);
                AddCell(ExcelBuffer, RowNo, 2, Format(PriceLine."Source Type"), false);
                AddCell(ExcelBuffer, RowNo, 3, PriceLine."Source No.", false);
                AddCell(ExcelBuffer, RowNo, 4, PriceLine."Asset No.", false);
                AddCell(ExcelBuffer, RowNo, 5, PriceLine."Currency Code", false);
                AddDecimalCell(ExcelBuffer, RowNo, 6, PriceLine."Minimum Quantity");
                AddDecimalCell(ExcelBuffer, RowNo, 7, PriceLine."Unit Price");

                RowNo += 1;
            until PriceLine.Next() = 0;

        ExcelBuffer.WriteSheet(SheetName, CompanyName, UserId);
    end;

    local procedure CreateCategorySheets(var ExcelBuffer: Record "Excel Buffer" temporary)
    var
        ItemCategory: Record "Item Category";
    begin
        if ItemCategory.FindSet() then
            repeat
                CreateCategorySheet(ExcelBuffer, ItemCategory.Code);
            until ItemCategory.Next() = 0;
    end;

    local procedure CreateCategorySheet(var ExcelBuffer: Record "Excel Buffer" temporary; CategoryCode: Code[20])
    var
        Item: Record Item;
        Setup: Record "LMS Price Export Setup";
        RowNo: Integer;
        ColNo: Integer;
    begin
        ExcelBuffer.SelectOrAddSheet(CategoryCode);

        AddCell(ExcelBuffer, 1, 1, 'Code', true);
        AddCell(ExcelBuffer, 1, 2, 'CATEGORY', true);

        ColNo := 3;

        Setup.SetCurrentKey("Sort Order");
        Setup.SetRange("Include In Matrix", true);

        if Setup.FindSet() then
            repeat
                AddCell(ExcelBuffer, 1, ColNo, Setup."Display Name", true);
                ColNo += 1;
            until Setup.Next() = 0;

        RowNo := 2;

        Item.SetRange("Item Category Code", CategoryCode);
        Item.SetRange(Blocked, false);

        if Item.FindSet() then
            repeat
                AddCell(ExcelBuffer, RowNo, 1, Item."No.", false);
                AddCell(ExcelBuffer, RowNo, 2, Item."Item Category Code", false);

                ColNo := 3;

                Setup.Reset();
                Setup.SetCurrentKey("Sort Order");
                Setup.SetRange("Include In Matrix", true);

                if Setup.FindSet() then
                    repeat
                        AddDecimalCell(
                            ExcelBuffer,
                            RowNo,
                            ColNo,
                            GetUnitPrice(Item."No.", Setup."Price List Code", Setup."Minimum Quantity")
                        );
                        ColNo += 1;
                    until Setup.Next() = 0;

                RowNo += 1;
            until Item.Next() = 0;

        ExcelBuffer.WriteSheet(CategoryCode, CompanyName, UserId);
    end;

    local procedure GetUnitPrice(ItemNo: Code[20]; PriceListCode: Code[20]; MinQty: Decimal): Decimal
    var
        PriceLine: Record "Price List Line";
    begin
        PriceLine.SetRange("Price List Code", PriceListCode);
        PriceLine.SetRange("Asset Type", PriceLine."Asset Type"::Item);
        PriceLine.SetRange("Asset No.", ItemNo);
        PriceLine.SetRange("Minimum Quantity", MinQty);

        if PriceLine.FindFirst() then
            exit(PriceLine."Unit Price");

        exit(0);
    end;

    local procedure AddCell(var ExcelBuffer: Record "Excel Buffer" temporary; RowNo: Integer; ColNo: Integer; Value: Text; IsBold: Boolean)
    begin
        ExcelBuffer.AddColumn(Value, false, '', IsBold, false, false, '', ExcelBuffer."Cell Type"::Text);
    end;

    local procedure AddDecimalCell(var ExcelBuffer: Record "Excel Buffer" temporary; RowNo: Integer; ColNo: Integer; Value: Decimal)
    begin
        ExcelBuffer.AddColumn(Value, false, '', false, false, false, '$#,##0.00', ExcelBuffer."Cell Type"::Number);
    end;
}