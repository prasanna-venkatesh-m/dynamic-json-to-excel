//using System;
//using System.Collections.Generic;
//using System.IO;
//using Newtonsoft.Json.Linq;
//using ClosedXML.Excel;

//class Program
//{
//    static void Main(string[] args)
//    {
//        string inputPath = @"C:\Users\PrasannaVenkateshM\Downloads\Surepass-Response-2 - Copy.xlsx";
//        string outputPath = @"C:\Users\PrasannaVenkateshM\Downloads\Output.xlsx";

//        using (var workbook = new XLWorkbook(inputPath))
//        {
//            var worksheet = workbook.Worksheet(1);
//            var rows = worksheet.RangeUsed().RowsUsed();

//            int lastColumn = worksheet.LastColumnUsed().ColumnNumber();
//            int newColumnStart = lastColumn + 1;

//            // Add headers
//            worksheet.Cell(1, newColumnStart).Value = "Full_Name";
//            worksheet.Cell(1, newColumnStart + 1).Value = "DOB";
//            worksheet.Cell(1, newColumnStart + 2).Value = "Gender";
//            worksheet.Cell(1, newColumnStart + 3).Value = "Age";

//            int rowIndex = 2;

//            foreach (var row in rows)
//            {
//                if (row.RowNumber() == 1)
//                    continue;

//                string apiResponse = row.Cell("C").GetString(); // APIResponse column

//                if (string.IsNullOrEmpty(apiResponse))
//                    continue;

//                apiResponse = apiResponse.Replace("\"\"", "\"");

//                try
//                {
//                    JObject json = JObject.Parse(apiResponse);

//                    var details = json["data"]?["details"];
//                    var personal = details?["personal_info"];
//                    var phones = details?["phone_info"];
//                    var addresses = details?["address_info"];

//                    int col = newColumnStart;

//                    // Personal Info
//                    worksheet.Cell(rowIndex, col++).Value = personal?["full_name"]?.ToString();
//                    worksheet.Cell(rowIndex, col++).Value = personal?["dob"]?.ToString();
//                    worksheet.Cell(rowIndex, col++).Value = personal?["gender"]?.ToString();
//                    worksheet.Cell(rowIndex, col++).Value = personal?["age"]?.ToString();

//                    // -------- All Phones --------
//                    if (phones != null)
//                    {
//                        int phoneIndex = 1;
//                        foreach (var phone in phones)
//                        {
//                            worksheet.Cell(1, col).Value = $"Phone_{phoneIndex}_Number";
//                            worksheet.Cell(rowIndex, col++).Value = phone["number"]?.ToString();

//                            worksheet.Cell(1, col).Value = $"Phone_{phoneIndex}_Type";
//                            worksheet.Cell(rowIndex, col++).Value = phone["type_code"]?.ToString();

//                            worksheet.Cell(1, col).Value = $"Phone_{phoneIndex}_Reported";
//                            worksheet.Cell(rowIndex, col++).Value = phone["reported_date"]?.ToString();

//                            phoneIndex++;
//                        }
//                    }

//                    // -------- All Addresses --------
//                    if (addresses != null)
//                    {
//                        int addressIndex = 1;
//                        foreach (var addr in addresses)
//                        {
//                            worksheet.Cell(1, col).Value = $"Address_{addressIndex}";
//                            worksheet.Cell(rowIndex, col++).Value = addr["address"]?.ToString();

//                            worksheet.Cell(1, col).Value = $"Address_{addressIndex}_State";
//                            worksheet.Cell(rowIndex, col++).Value = addr["state"]?.ToString();

//                            worksheet.Cell(1, col).Value = $"Address_{addressIndex}_Postal";
//                            worksheet.Cell(rowIndex, col++).Value = addr["postal"]?.ToString();

//                            worksheet.Cell(1, col).Value = $"Address_{addressIndex}_Type";
//                            worksheet.Cell(rowIndex, col++).Value = addr["type"]?.ToString();

//                            worksheet.Cell(1, col).Value = $"Address_{addressIndex}_Reported";
//                            worksheet.Cell(rowIndex, col++).Value = addr["reported_date"]?.ToString();

//                            addressIndex++;
//                        }
//                    }
//                }
//                catch (Exception ex)
//                {
//                    Console.WriteLine($"Error at row {rowIndex}: {ex.Message}");
//                }

//                rowIndex++;
//            }

//            workbook.SaveAs(outputPath);
//        }

//        Console.WriteLine("✅ Conversion Completed!");
//    }
//}

using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        string inputPath = @"C:\Users\PrasannaVenkateshM\Downloads\Surepass-Response-2 - Copy.xlsx";
        string outputPath = @"C:\Users\PrasannaVenkateshM\Downloads\Output-2.xlsx";


        var records = new List<JObject>();

        int maxPhones = 0;
        int maxAddresses = 0;
        int maxEmails = 0;
        int maxPan = 0;
        int maxAadhaar = 0;
        int maxVoter = 0;
        int maxOtherId = 0;

        // ---------------- FIRST PASS ----------------
        using (var wb = new XLWorkbook(inputPath))
        {
            var ws = wb.Worksheet(1);

            foreach (var row in ws.RangeUsed().RowsUsed().Skip(1))
            {
                string jsonText = row.Cell(3).GetString();
                if (string.IsNullOrWhiteSpace(jsonText)) continue;

                jsonText = jsonText.Replace("\"\"", "\"");

                var json = JObject.Parse(jsonText);
                records.Add(json);

                var details = json["data"]?["details"];

                maxPhones = Math.Max(maxPhones, details?["phone_info"]?.Count() ?? 0);
                maxAddresses = Math.Max(maxAddresses, details?["address_info"]?.Count() ?? 0);
                maxEmails = Math.Max(maxEmails, details?["email_info"]?.Count() ?? 0);
                maxPan = Math.Max(maxPan, details?["identity_info"]?["pan_number"]?.Count() ?? 0);
                maxAadhaar = Math.Max(maxAadhaar, details?["identity_info"]?["aadhaar_number"]?.Count() ?? 0);
                maxVoter = Math.Max(maxVoter, details?["identity_info"]?["voter_id"]?.Count() ?? 0);
                maxOtherId = Math.Max(maxOtherId, details?["identity_info"]?["other_id"]?.Count() ?? 0);
            }
        }

        // ---------------- CREATE OUTPUT ----------------
        using (var newWb = new XLWorkbook())
        {
            var ws = newWb.AddWorksheet("Output");

            int col = 1;

            // Fixed fields
            ws.Cell(1, col++).Value = "Full_Name";
            ws.Cell(1, col++).Value = "DOB";
            ws.Cell(1, col++).Value = "Gender";
            ws.Cell(1, col++).Value = "Age";
            ws.Cell(1, col++).Value = "Total_Income";
            ws.Cell(1, col++).Value = "Occupation";

            // Dynamic Phones
            for (int i = 1; i <= maxPhones; i++)
            {
                ws.Cell(1, col++).Value = $"Phone_{i}_Number";
                ws.Cell(1, col++).Value = $"Phone_{i}_Type";
                ws.Cell(1, col++).Value = $"Phone_{i}_Reported";
            }

            // Dynamic Addresses
            for (int i = 1; i <= maxAddresses; i++)
            {
                ws.Cell(1, col++).Value = $"Address_{i}";
                ws.Cell(1, col++).Value = $"Address_{i}_State";
                ws.Cell(1, col++).Value = $"Address_{i}_Postal";
                ws.Cell(1, col++).Value = $"Address_{i}_Type";
                ws.Cell(1, col++).Value = $"Address_{i}_Reported";
            }

            // Emails
            for (int i = 1; i <= maxEmails; i++)
                ws.Cell(1, col++).Value = $"Email_{i}";

            // Identity
            for (int i = 1; i <= maxPan; i++)
                ws.Cell(1, col++).Value = $"PAN_{i}";

            for (int i = 1; i <= maxAadhaar; i++)
                ws.Cell(1, col++).Value = $"Aadhaar_{i}";

            for (int i = 1; i <= maxVoter; i++)
                ws.Cell(1, col++).Value = $"VoterID_{i}";

            for (int i = 1; i <= maxOtherId; i++)
                ws.Cell(1, col++).Value = $"OtherID_{i}";

            // ---------------- FILL DATA ----------------
            int rowIndex = 2;

            foreach (var json in records)
            {
                col = 1;

                var details = json["data"]?["details"];
                var personal = details?["personal_info"];

                ws.Cell(rowIndex, col++).Value = personal?["full_name"]?.ToString();
                ws.Cell(rowIndex, col++).Value = personal?["dob"]?.ToString();
                ws.Cell(rowIndex, col++).Value = personal?["gender"]?.ToString();
                ws.Cell(rowIndex, col++).Value = personal?["age"]?.ToString();
                ws.Cell(rowIndex, col++).Value = personal?["total_income"]?.ToString();
                ws.Cell(rowIndex, col++).Value = personal?["occupation"]?.ToString();

                // Phones
                var phones = details?["phone_info"];
                for (int i = 0; i < maxPhones; i++)
                {
                    if (phones != null && i < phones.Count())
                    {
                        ws.Cell(rowIndex, col++).Value = phones[i]?["number"]?.ToString();
                        ws.Cell(rowIndex, col++).Value = phones[i]?["type_code"]?.ToString();
                        ws.Cell(rowIndex, col++).Value = phones[i]?["reported_date"]?.ToString();
                    }
                    else col += 3;
                }

                // Addresses
                var addresses = details?["address_info"];
                for (int i = 0; i < maxAddresses; i++)
                {
                    if (addresses != null && i < addresses.Count())
                    {
                        ws.Cell(rowIndex, col++).Value = addresses[i]?["address"]?.ToString();
                        ws.Cell(rowIndex, col++).Value = addresses[i]?["state"]?.ToString();
                        ws.Cell(rowIndex, col++).Value = addresses[i]?["postal"]?.ToString();
                        ws.Cell(rowIndex, col++).Value = addresses[i]?["type"]?.ToString();
                        ws.Cell(rowIndex, col++).Value = addresses[i]?["reported_date"]?.ToString();
                    }
                    else col += 5;
                }

                rowIndex++;
            }

            newWb.SaveAs(outputPath);
        }

        Console.WriteLine("✅ Fully Dynamic Extraction Completed");
    }
}

