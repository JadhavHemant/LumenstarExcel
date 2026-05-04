pageextension 50100 "LMS Item List Ext" extends "Item List"
{
    actions
    {
        addlast(Processing)
        {
            action(GenerateLumenstarPriceList)
            {
                Caption = 'Generate Price List Workbook';
                ApplicationArea = All;
                Image = ExportToExcel;

                trigger OnAction()
                var
                    ExportMgt: Codeunit "LMS Price List Export Mgt.";
                begin
                    ExportMgt.ExportWorkbook();
                end;
            }
        }
    }
}