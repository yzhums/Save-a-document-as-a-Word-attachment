pageextension 50115 PostedSalesInvoicesExt extends "Posted Sales Invoices"
{
    actions
    {
        addafter(AttachAsPDF_Promoted)
        {
            actionref(AttachAsWord_Promoted; AttachAsWord)
            {
            }
        }
        addafter(AttachAsPDF)
        {
            action(AttachAsWord)
            {
                ApplicationArea = All;
                Caption = 'Attach as Word';
                Image = PrintAttachment;
                ToolTip = 'Create a Word file and attach it to the document.';

                trigger OnAction()
                var
                    ReportSelection: Record "Report Selections";
                    SalesInvHeader: Record "Sales Invoice Header";
                    SalesInvHeader2: Record "Sales Invoice Header";
                    TempReportSelections: Record "Report Selections" temporary;
                    TempBlob: Codeunit "Temp Blob";
                    FileName: Text[50];
                    InS: InStream;
                    DocAttach: Record "Document Attachment";
                    ReportDistributionMgt: Codeunit "Report Distribution Management";
                    ReportCaption: Text[250];
                    DocumentLanguageCode: Code[10];
                begin
                    SalesInvHeader.Reset();
                    CurrPage.SetSelectionFilter(SalesInvHeader);
                    if SalesInvHeader.FindSet() then
                        repeat
                            SalesInvHeader2.Get(SalesInvHeader."No.");
                            SalesInvHeader2.SetRecFilter();
                            ReportSelection.FindReportUsageForCust(Enum::"Report Selection Usage"::"S.Invoice", SalesInvHeader2."Bill-to Customer No.", TempReportSelections);
                            Clear(TempBlob);
                            SaveReportAsPDFInTempBlob(TempBlob, TempReportSelections."Report ID", SalesInvHeader2, TempReportSelections."Custom Report Layout Code", Enum::"Report Selection Usage"::"S.Invoice", TempReportSelections);
                            TempBlob.CreateInStream(InS);
                            DocAttach.Init();
                            DocAttach.Validate("Table ID", SalesInvHeader2.RecordId.TableNo);
                            DocAttach.Validate("No.", SalesInvHeader2."No.");
                            ReportCaption := ReportDistributionMgt.GetReportCaption(TempReportSelections."Report ID", DocumentLanguageCode);
                            FileName := StrSubstNo('%1 %2 %3', TempReportSelections."Report ID", ReportCaption, SalesInvHeader2."No.");
                            DocAttach.Validate("File Name", FileName);
                            DocAttach.Validate("File Extension", 'docx');
                            DocAttach."Document Reference ID".ImportStream(InS, FileName);
                            DocAttach.Insert(true);
                        until SalesInvHeader.Next() = 0;
                end;
            }
        }
    }

    local procedure SaveReportAsPDFInTempBlob(var TempBlob: Codeunit "Temp Blob"; ReportID: Integer; RecordVariant: Variant; LayoutCode: Code[20]; ReportUsage: Enum "Report Selection Usage"; TempReportSelections: Record "Report Selections" temporary)
    var
        ReportLayoutSelectionLocal: Record "Report Layout Selection";
        CustomLayoutReporting: Codeunit "Custom Layout Reporting";
        CustomerStatementSub: codeunit "Customer Statement Subscr";
        LastUsedParameters: Text;
        IsHandled: Boolean;
        OutStream: OutStream;
    begin
        if TempReportSelections."Report Layout Name" <> '' then
            ReportLayoutSelectionLocal.SetTempLayoutSelectedName(TempReportSelections."Report Layout Name", TempReportSelections."Report Layout AppID")
        else
            ReportLayoutSelectionLocal.SetTempLayoutSelected(LayoutCode);
        BindSubscription(CustomerStatementSub);
        UnbindSubscription(CustomerStatementSub);

        if not IsHandled then begin
            TempBlob.CreateOutStream(OutStream);
            LastUsedParameters := CustomLayoutReporting.GetReportRequestPageParameters(ReportID);
            Report.SaveAs(ReportID, LastUsedParameters, ReportFormat::Word, OutStream, GetRecRef(RecordVariant));
        end;

        ReportLayoutSelectionLocal.ClearTempLayoutSelected();

        Commit();
    end;

    local procedure GetRecRef(RecVariant: Variant) RecRef: RecordRef
    begin
        if RecVariant.IsRecordRef() then
            exit(RecVariant);
        if RecVariant.IsRecord() then
            RecRef.GetTable(RecVariant);
    end;
}
