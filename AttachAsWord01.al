codeunit 50113 SaveAsWordHandler
{
    [EventSubscriber(ObjectType::Table, Database::"Report Selections", OnBeforeSaveReportAsPDF, '', false, false)]
    local procedure ReportSelections_OnBeforeSaveReportAsPDF(var ReportID: Integer; RecordVariant: Variant; var IsHandled: Boolean; var TempBlob: Codeunit "Temp Blob")
    var
        OutStream: OutStream;
        LastUsedParameters: Text;
        CustomLayoutReporting: Codeunit "Custom Layout Reporting";
    begin
        if ReportID = Report::"Standard Sales - Invoice" then begin
            TempBlob.CreateOutStream(OutStream);
            LastUsedParameters := CustomLayoutReporting.GetReportRequestPageParameters(ReportID);
            Report.SaveAs(ReportID, LastUsedParameters, ReportFormat::Word, OutStream, GetRecRef(RecordVariant));
            IsHandled := true;
        end;
    end;

    local procedure GetRecRef(RecVariant: Variant) RecRef: RecordRef
    begin
        if RecVariant.IsRecordRef() then
            exit(RecVariant);
        if RecVariant.IsRecord() then
            RecRef.GetTable(RecVariant);
    end;

    [EventSubscriber(ObjectType::Table, Database::"Document Attachment", OnBeforeSaveAttachment, '', false, false)]
    local procedure DocumentAttachment_OnBeforeSaveAttachment(var RecRef: RecordRef; var FileName: Text)
    var
        FileManagement: Codeunit "File Management";
        FileNameWithoutExtension: Text;
    begin
        if RecRef.RecordId.TableNo = Database::"Sales Invoice Header" then begin
            FileNameWithoutExtension := FileManagement.GetFileNameWithoutExtension(FileName);
            FileName := FileNameWithoutExtension + '.docx';
        end;
    end;
}
