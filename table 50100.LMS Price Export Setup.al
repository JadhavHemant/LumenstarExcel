table 50100 "LMS Price Export Setup"
{
    DataClassification = CustomerContent;

    fields
    {
        field(1; "Code"; Code[20]) { }
        field(2; "Price List Code"; Code[20]) { }
        field(3; "Assign-to Type"; Enum "Price Source Type") { }
        field(4; "Assign-to No."; Code[20]) { }
        field(5; "Display Name"; Text[50]) { }
        field(6; "Minimum Quantity"; Decimal) { }
        field(7; "Currency Code"; Code[10]) { }
        field(8; "Include In Matrix"; Boolean) { }
        field(9; "Sort Order"; Integer) { }
    }

    keys
    {
        key(PK; "Code") { Clustered = true; }
        key(Sort; "Sort Order") { }
    }
}