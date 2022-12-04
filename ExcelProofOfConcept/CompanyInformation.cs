using System;

namespace ExcelProofOfConcept;

public record CompanyInformation(string CompanyName)
{
    public string CompanyName { get; set; } = CompanyName;
}

public record AuditItem(string ClientName, string FileName, string Exists, string Notes)
{
    public string Exists { get; set; } = Exists;
public string Notes { get; set; } = Notes;
};
