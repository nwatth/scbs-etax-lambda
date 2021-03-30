using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace Report
{
    class Program
    {
        static async Task Main(string[] args)
        {

//            var json = @"{
//   ""documents"":[
//      {
//         ""pages"":[
//            {
//               ""payload"":{
//                  ""amount"":""399,444.91"",
//                  ""atsFee"":"""",
//                  ""remark"":""***** บริษัทจะหักเงินฝากของท่านจาก บัญชี \""บล.ไทยพาณิชย์ เพื่อลูกค้า\"" เพื่อมาชำระรายการนี้ *****"",
//                  ""trades"":[
//                     {
//                        ""vat"":""1.07"",
//                        ""unit"":""5,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""15.32"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 159935"",
//                        ""totalAmount"":""9,766.39"",
//                        ""contractNumber"":""BU-69986""
//                     },
//                     {
//                        ""vat"":""2.14"",
//                        ""unit"":""10,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""30.61"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 160062"",
//                        ""totalAmount"":""19,532.75"",
//                        ""contractNumber"":""BU-70134""
//                     },
//                     {
//                        ""vat"":""1.07"",
//                        ""unit"":""5,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""15.32"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 160331"",
//                        ""totalAmount"":""9,766.39"",
//                        ""contractNumber"":""BU-70325""
//                     },
//                     {
//                        ""vat"":""4.29"",
//                        ""unit"":""20,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""61.23"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 160457"",
//                        ""totalAmount"":""39,065.52"",
//                        ""contractNumber"":""BU-70418""
//                     },
//                     {
//                        ""vat"":""1.07"",
//                        ""unit"":""5,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""15.32"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 155142"",
//                        ""totalAmount"":""9,766.39"",
//                        ""contractNumber"":""BU-72908""
//                     },
//                     {
//                        ""vat"":""1.07"",
//                        ""unit"":""5,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""15.32"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 154853"",
//                        ""totalAmount"":""9,766.39"",
//                        ""contractNumber"":""BU-77292""
//                     },
//                     {
//                        ""vat"":""2.14"",
//                        ""unit"":""10,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""30.62"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 154365"",
//                        ""totalAmount"":""19,532.76"",
//                        ""contractNumber"":""BU-77293""
//                     },
//                     {
//                        ""vat"":""4.29"",
//                        ""unit"":""20,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""61.25"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 154204"",
//                        ""totalAmount"":""39,065.54"",
//                        ""contractNumber"":""BU-79906""
//                     },
//                     {
//                        ""vat"":""1.07"",
//                        ""unit"":""5,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""15.32"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 169297"",
//                        ""totalAmount"":""9,766.39"",
//                        ""contractNumber"":""BU-81997""
//                     },
//                     {
//                        ""vat"":""2.14"",
//                        ""unit"":""10,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""30.63"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 169577"",
//                        ""totalAmount"":""19,532.77"",
//                        ""contractNumber"":""BU-81998""
//                     },
//                     {
//                        ""vat"":""1.07"",
//                        ""unit"":""5,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""15.32"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 169749"",
//                        ""totalAmount"":""9,766.39"",
//                        ""contractNumber"":""BU-82331""
//                     },
//                     {
//                        ""vat"":""4.29"",
//                        ""unit"":""20,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""61.23"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 155892"",
//                        ""totalAmount"":""39,065.52"",
//                        ""contractNumber"":""BU-82333""
//                     },
//                     {
//                        ""vat"":""4.32"",
//                        ""unit"":""20,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""61.23"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 178371"",
//                        ""totalAmount"":""39,065.55"",
//                        ""contractNumber"":""BU-86119""
//                     },
//                     {
//                        ""vat"":""4.28"",
//                        ""unit"":""20,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""61.08"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 178562"",
//                        ""totalAmount"":""39,065.36"",
//                        ""contractNumber"":""BU-87139""
//                     },
//                     {
//                        ""vat"":""2.14"",
//                        ""unit"":""10,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""30.62"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 179062"",
//                        ""totalAmount"":""19,532.76"",
//                        ""contractNumber"":""BU-87140""
//                     },
//                     {
//                        ""vat"":""1.07"",
//                        ""unit"":""5,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""15.32"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 179254"",
//                        ""totalAmount"":""9,766.39"",
//                        ""contractNumber"":""BU-87141""
//                     },
//                     {
//                        ""vat"":""1.07"",
//                        ""unit"":""5,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""15.32"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 179431"",
//                        ""totalAmount"":""9,766.39"",
//                        ""contractNumber"":""BU-87142""
//                     },
//                     {
//                        ""vat"":""2.14"",
//                        ""unit"":""10,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""30.62"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 181080"",
//                        ""totalAmount"":""19,532.76"",
//                        ""contractNumber"":""BU-87956""
//                     },
//                     {
//                        ""vat"":""2.14"",
//                        ""unit"":""10,000"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""30.62"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 181327"",
//                        ""totalAmount"":""19,532.76"",
//                        ""contractNumber"":""BU-92958""
//                     },
//                     {
//                        ""vat"":""0.96"",
//                        ""unit"":""4,500"",
//                        ""unitPrice"":""1.9500"",
//                        ""commission"":""13.78"",
//                        ""securities"":""SCM"",
//                        ""orderNumber"":""I 182397"",
//                        ""totalAmount"":""8,789.74"",
//                        ""contractNumber"":""BU-92959""
//                     }
//                  ],
//                  ""barcode"":""By_NoPrint_SDC_902005248"",
//                  ""channel"":"""",
//                  ""taxRate"":"""",
//                  ""vatRate"":""ภาษีมูลค่าเพิ่มอัตราร้อยละ  7.00"",
//                  ""otherFee"":"""",
//                  ""totalFee"":""ค่าธรรมเนียมรวม"",
//                  ""atsFeeVat"":"""",
//                  ""buyAmount"":""399,444.91"",
//                  ""payAmount"":""399,444.91"",
//                  ""tradeSize"":20,
//                  ""branchCode"":""00000 :"",
//                  ""pageNumber"":""1/1"",
//                  ""sellAmount"":""0.00"",
//                  ""summaryVat"":""43.83"",
//                  ""accountType"":""Cash Balance"",
//                  ""commission1"":""Commission"",
//                  ""commission2"":""="",
//                  ""commission3"":""598.16"",
//                  ""tradingDate"":""22/01/2021"",
//                  ""tradingFee1"":""Trading and Regulatory Fee"",
//                  ""tradingFee2"":""="",
//                  ""tradingFee3"":""23.93"",
//                  ""atsFeeAmount"":"""",
//                  ""clearingFee1"":""Clearing Fee"",
//                  ""clearingFee2"":""="",
//                  ""clearingFee3"":""3.99"",
//                  ""accountNumber"":""90-2-005248"",
//                  ""customerTaxId"":""3209600126634"",
//                  ""payAmountWord"":""สามแสนเก้าหมื่นเก้าพันสี่ร้อยสี่สิบสี่บาทเก้าสิบเอ็ดสตางค์"",
//                  ""receiveAmount"":"""",
//                  ""accountOfficer"":""7500 E-Business"",
//                  ""documentNumber"":""DN-20210122-11799"",
//                  ""settlementDate"":""26/01/2021"",
//                  ""atsFeeCommission"":"""",
//                  ""customerFullname"":""นาย ชื่อ90005248 นามสกุล90005248"",
//                  ""taxInvoiceNumber"":""2021-00430505"",
//                  ""customerWorkplace"":"""",
//                  ""receiveAmountWord"":"""",
//                  ""summaryCommission"":""626.08"",
//                  ""branchForTaxInvoice"":""00000 : สำนักงานใหญ่"",
//                  ""customerAddressLine1"":""118/9 หมู่10"",
//                  ""customerAddressLine2"":""ต.หนองปรือ"",
//                  ""customerAddressLine3"":""อ.บางละมุง จ.ชลบุรี 20150"",
//                  ""numberCustomerCodeAndDocumentMessageRoute"":""C0-528  AA""
//               },
//               ""templateId"":""SCBS_DLC_ONE_PAGE_20210101""
//            }
//         ],
//         ""documentConfig"":{
//            ""protect"":""01071977"",
//            ""signing"":true,
//            ""fileName"":""SCBSDLC_20210122_90-2-005248_00.pdf"",
//            ""sendMail"":true,
//            ""archivePaths"":[
//               {
//                  ""path"":[
//                     ""Daily_2021/STOCK/SDC/01_2021/20210122/PDF/NoPrint/""
//                  ],
//                  ""bucketName"":""scbs-s3-etax-stroage-uat""
//               }
//            ]
//         }
//      },
//      {
//         ""pages"":[
//            {
//               ""payload"":{
//                  ""date"":""22/01/2021"",
//                  ""branch"":""(บล.ไทยพาณิชย์ จก. สำนักงานใหญ่)"",
//                  ""remark"":""ฝาก Nsips (eService)"",
//                  ""userId"":""(APIBPM)"",
//                  ""barcode"":""By_NoPrint_SDC_902005248"",
//                  ""program"":""KPR006/902005248"",
//                  ""cashAmount"":""294,000.00"",
//                  ""pageNumber"":""1/1"",
//                  ""cashDuedate"":""22/01/2021"",
//                  ""accountNumber"":""902005248             "",
//                  ""chequeAmount1"":"""",
//                  ""chequeAmount2"":"""",
//                  ""chequeAmount3"":"""",
//                  ""chequeNumber1"":"""",
//                  ""chequeNumber2"":"""",
//                  ""chequeNumber3"":"""",
//                  ""chequeDuedate1"":"""",
//                  ""chequeDuedate2"":"""",
//                  ""chequeDuedate3"":"""",
//                  ""transferByCash"":""x"",
//                  ""amountTotalWord"":""(สองแสนเก้าหมื่นสี่พันบาทถ้วน)"",
//                  ""chequeBankName1"":"""",
//                  ""chequeBankName2"":"""",
//                  ""chequeBankName3"":"""",
//                  ""referenceNumber"":""DH-20210122-C0155"",
//                  ""customerFullname"":""นาย ชื่อ90005248 นามสกุล90005248"",
//                  ""transferByCheque"":"""",
//                  ""chequeBankBranch1"":"""",
//                  ""chequeBankBranch2"":"""",
//                  ""chequeBankBranch3"":"""",
//                  ""customerAddressLine1"":""118/9 หมู่10"",
//                  ""customerAddressLine2"":""ต.หนองปรือ"",
//                  ""amountTotalAndInternet"":""294,000.00"",
//                  ""customerAddressLine3And4"":""อ.บางละมุง จ.ชลบุรี 20150""
//               },
//               ""templateId"":""SCBS_DEP_20210101""
//            }
//         ],
//         ""documentConfig"":{
//            ""protect"":""01071977"",
//            ""signing"":true,
//            ""fileName"":""SCBS_DEP_20210122_90-2-005248_00.pdf"",
//            ""sendMail"":true,
//            ""archivePaths"":[
//               {
//                  ""path"":[
//                     ""Daily_2021/STOCK/SDC/01_2021/20210122/PDF/NoPrint/""
//                  ],
//                  ""bucketName"":""scbs-s3-etax-stroage-uat""
//               }
//            ]
//         }
//      },
//      {
//         ""pages"":[
//            {
//               ""payload"":{
//                  ""date"":""22/01/2021"",
//                  ""branch"":""(บล.ไทยพาณิชย์ จก. สำนักงานใหญ่)"",
//                  ""remark"":""ฝาก Nsips (eService)"",
//                  ""userId"":""(APIBPM)"",
//                  ""barcode"":""By_NoPrint_SDC_902005248"",
//                  ""program"":""KPR006/902005248"",
//                  ""cashAmount"":""190,000.00"",
//                  ""pageNumber"":""1/1"",
//                  ""cashDuedate"":""22/01/2021"",
//                  ""accountNumber"":""902005248             "",
//                  ""chequeAmount1"":"""",
//                  ""chequeAmount2"":"""",
//                  ""chequeAmount3"":"""",
//                  ""chequeNumber1"":"""",
//                  ""chequeNumber2"":"""",
//                  ""chequeNumber3"":"""",
//                  ""chequeDuedate1"":"""",
//                  ""chequeDuedate2"":"""",
//                  ""chequeDuedate3"":"""",
//                  ""transferByCash"":""x"",
//                  ""amountTotalWord"":""(หนึ่งแสนเก้าหมื่นบาทถ้วน)"",
//                  ""chequeBankName1"":"""",
//                  ""chequeBankName2"":"""",
//                  ""chequeBankName3"":"""",
//                  ""referenceNumber"":""DH-20210122-C016W"",
//                  ""customerFullname"":""นาย ชื่อ90005248 นามสกุล90005248"",
//                  ""transferByCheque"":"""",
//                  ""chequeBankBranch1"":"""",
//                  ""chequeBankBranch2"":"""",
//                  ""chequeBankBranch3"":"""",
//                  ""customerAddressLine1"":""118/9 หมู่10"",
//                  ""customerAddressLine2"":""ต.หนองปรือ"",
//                  ""amountTotalAndInternet"":""190,000.00"",
//                  ""customerAddressLine3And4"":""อ.บางละมุง จ.ชลบุรี 20150""
//               },
//               ""templateId"":""SCBS_DEP_20210101""
//            }
//         ],
//         ""documentConfig"":{
//            ""protect"":""01071977"",
//            ""signing"":true,
//            ""fileName"":""SCBS_DEP_20210122_90-2-005248_01.pdf"",
//            ""sendMail"":true,
//            ""archivePaths"":[
//               {
//                  ""path"":[
//                     ""Daily_2021/STOCK/SDC/01_2021/20210122/PDF/NoPrint/""
//                  ],
//                  ""bucketName"":""scbs-s3-etax-stroage-uat""
//               }
//            ]
//         }
//      }
//   ],
//   ""documentMeta"":{
//      ""product"":""EQUITY"",
//      ""dataDate"":""2021-01-22"",
//      ""sendMail"":true,
//      ""totalFile"":3,
//      ""deliveryInfo"":{
//         ""birthday"":""01071977"",
//         ""customerEmail"":""scbs.uat.etax.tks@@scb.co.th"",
//         ""customerFullName"":""นาย ชื่อ90005248 นามสกุล90005248"",
//         ""accountNumberFormat"":""90-2-005248""
//      },
//      ""documentType"":""SDC"",
//      ""accountNumber"":""902005248"",
//      ""marketingInfo"":{
//         ""email"":""scbsonline@scb.co.th"",
//         ""teamId"":""020108"",
//         ""teamName"":""ฝ่ายธุรกิจอิเล็คทรอนิกส์"",
//         ""marketingId"":""7500"",
//         ""marketingName"":""ฝ่าย E-Business (7500)"",
//         ""officeTelNumber"":""029491234""
//      },
//      ""printDelivery"":""NoPrint"",
//      ""jobExecutionId"":3
//   }
//}";

//            var data = JsonDocument.Parse(json).RootElement;
//            foreach (var item in data.GetProperty("documents").EnumerateArray())
//            {
//                var doc = item.GetProperty("documentConfig");
//                var file = doc.GetProperty("fileName").GetString();

//                foreach (var archive in doc.GetProperty("archivePaths").EnumerateArray())
//                {
                    
//                    var bucket = archive.GetProperty("bucketName").GetString();
//                    foreach (var path in archive.GetProperty("path").EnumerateArray())
//                    {
//                        var path1 = path.GetString() + file;
//                    }
//                }
                
//            } 

            
            var dataBounce = await LoadBounce(2020);
            var branch = LoadBranch();

            ReportBounce(dataBounce, branch);

            var dataSummary = await LoadSummary(2020);
            ReportSummary(dataSummary);
            Console.WriteLine("Hello World!");  
        }

        private static async Task<bool> CheckDuplicate(string msgId)
        {
            var sql = $@"SELECT 1 FROM public.job_step_execution WHERE message_id = '{msgId}' AND status = 'SUCCESS' LIMIT 1";
            using (var conn = await Utility.CreateConnection())
            {
                var result = await conn.ExecuteScalarAsync(sql);
                return result != null && (int)result != 0;
            }
        }

        public static void ReportBounce(List<Bounce> data, List<(string branch, string team)> branch)
        {
            var sep = Path.DirectorySeparatorChar;
            var excel = new FileInfo($"{Environment.CurrentDirectory}{sep}bounce.xlsx");

            using (var excelPackage = new ExcelPackage(excel))
            {
                var sheet = 0;
                var cells = excelPackage.Workbook.Worksheets[sheet].Cells;
                int i = 1;

                foreach (var item in data)
                {
                    i++;
                    cells["A" + i].Value = item.Date;
                    cells["B" + i].Value = item.Account;
                    cells["C" + i].Value = item.Team;
                    cells["D" + i].Value = branch.Where(w=>w.team.Contains(item.Team)).FirstOrDefault().branch;
                    cells["E" + i].Value = item.Name;
                    cells["F" + i].Value = item.Email;
                    cells["G" + i].Value = item.Remark;
                    cells["H" + i].Value = item.DocType;
                    cells["I" + i].Value = item.AoId;
                    cells["J" + i].Value = item.AoName;

                }

                if (i != 1)
                {
                    var range = $"A2:J{i}";
                    cells[range].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cells[range].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cells[range].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cells[range].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                }

                sheet++;


                foreach (var b in branch)
                {
                    i = 1;
                    cells = excelPackage.Workbook.Worksheets[sheet].Cells;

                    foreach (var item in data.Where(w=>b.team.Contains(w.Team)))
                    {
                        i++;
                        cells["A" + i].Value = item.Date;
                        cells["B" + i].Value = item.Account;
                        cells["C" + i].Value = item.Name;
                        cells["D" + i].Value = item.Email;
                        cells["E" + i].Value = item.Remark;
                        cells["F" + i].Value = item.DocType;
                        cells["G" + i].Value = item.AoId;
                        cells["H" + i].Value = item.AoName;
                        
                    }
                    if(i != 1)
                    {
                        var range = $"A2:H{i}";
                        cells[range].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        cells[range].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        cells[range].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        cells[range].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }
                    

                    sheet++;
                }

                excelPackage.SaveAs(new FileInfo($"{Environment.CurrentDirectory}{sep}report_bounce.xlsx"));
            }
        }

        public static void ReportSummary(Dictionary<DateTime, Summary> data)
        {

            var sep = Path.DirectorySeparatorChar;
            var excel = new FileInfo($"{Environment.CurrentDirectory}{sep}summary.xlsx");

            using (var excelPackage = new ExcelPackage(excel))
            {
                var colorSuccess = System.Drawing.Color.FromArgb(204, 153, 255);
                var colorFail = System.Drawing.Color.FromArgb(244, 176, 132);

                int i = 5;
                var cells = excelPackage.Workbook.Worksheets[0].Cells;
                var year = data.Count != 0 ? data.First().Key.Year : DateTime.Now.Year;
                DateTime start = new DateTime(year, 1, 1);
                DateTime end = new DateTime(year + 1, 1, 1);
                var month = 1;
                do
                {
                    if (month != start.Month)
                    {
                        month++;

                        if (month == 3 && year % 4 != 0)
                        {
                            cells["A67"].Value = "";
                            i += 4;
                        }
                        else
                        {
                            i += 3;
                        }
        
                        

                    }
                    var fillColor = false;

                    cells["A" + i].Value = start.ToString("dd/MM/yyyy");
                    if (data.ContainsKey(start))
                    {

                        if (data[start].DocType == "SDC")
                        {
                            fillColor = true;
                            
                            cells["E" + i].Style.Fill.BackgroundColor.SetColor(colorSuccess);
                            cells["E" + i].Style.Fill.BackgroundColor.Tint = 0;

                            cells["F" + i].Value = data[start].Success;
                            cells["F" + i].Style.Fill.BackgroundColor.SetColor(colorSuccess);
                            cells["F" + i].Style.Fill.BackgroundColor.Tint = 0;

                            cells["G" + i].Value = data[start].Fail;
                            cells["G" + i].Style.Fill.BackgroundColor.Tint = 0;
                            cells["G" + i].Style.Fill.BackgroundColor.SetColor(colorFail);
                        }
                    }

                    if (fillColor)
                    {

                        cells["B" + i].Style.Fill.BackgroundColor.SetColor(colorSuccess);
                        cells["B" + i].Style.Fill.BackgroundColor.Tint = 0;
                        cells["C" + i].Style.Fill.BackgroundColor.SetColor(colorSuccess);
                        cells["C" + i].Style.Fill.BackgroundColor.Tint = 0;
                        cells["D" + i].Style.Fill.BackgroundColor.SetColor(colorFail);
                        cells["D" + i].Style.Fill.BackgroundColor.Tint = 0;
                    }

                    i++;
                    start = start.AddDays(1);
                }
                while (start < end);

                excelPackage.SaveAs(new FileInfo($"{Environment.CurrentDirectory}{sep}report_summary.xlsx"));
            }
        }

        public static async Task<List<Bounce>> LoadBounce(int year)
        {
            var result = new List<Bounce>();
            var start = new DateTime(year, 1, 1).ToString("MM-dd-yyyy");
            var end = new DateTime(year, 12, 31).ToString("MM-dd-yyyy");

            var sql = $@"WITH b as ( 
                                      SELECT 
                                      data_date,
                                      account,
                                      document_type,
                                      remark
                                      FROM public.job_step_execution
                                      WHERE data_date >= '{start}' AND data_date <= '{end}' AND step_name = 'SentEmail' AND status = 'BOUNCE'
                                      ), 
                                      s as (
  	                                    SELECT 
                                        data_date,
                                        account,
                                        document_type,
                                        message_payload
                                        FROM public.job_step_execution
                                        WHERE data_date >= '{start}' AND data_date <= '{end}' AND step_name = 'SendingEmail' AND account IN (SELECT b.account FROM b)
                                      )
  
                                    SELECT 
                                      b.data_date,
                                      b.document_type,
                                      b.remark,
                                      s.message_payload#>>'{{documentMeta,deliveryInfo,accountNumberFormat}}' as account,
                                      s.message_payload#>>'{{documentMeta,deliveryInfo,customerFullname}}' as name,
                                      s.message_payload#>>'{{documentMeta,deliveryInfo,customerEmail}}' as email,
                                      s.message_payload#>>'{{documentMeta,marketingInfo,marketingName}}' as mkt,
                                      s.message_payload#>>'{{documentMeta,marketingInfo,marketingId}}' as mkt_id,
                                      s.message_payload#>>'{{documentMeta,marketingInfo,teamId}}' as team_id
                                    FROM b
                                    LEFT JOIN s ON b.account = s.account AND b.data_date = s.data_date  AND b.document_type = s.document_type";
            using (var connection = await Utility.CreateConnection())
            using (var reader = await connection.ExecuteReaderAsync(sql))
            {
                while (reader.Read())
                {
                    result.Add(new Bounce
                    {
                        Date = reader.GetDateTime(0).ToString("dd/MM/yyyy"),
                        DocType = reader.GetString(1),
                        Remark = reader.GetString(2),
                        Account = reader.GetString(3),
                        Name = reader.GetString(4),
                        Email = reader.GetString(5),
                        AoName = reader.GetString(6),
                        AoId = reader.GetString(7),
                        Team = reader.GetString(8),
                    });
                }
            }
            return result;
        }

        public static async Task<Dictionary<DateTime,Summary>> LoadSummary(int year)
        {
            var result = new Dictionary<DateTime, Summary>();
            var start = new DateTime(year, 1, 1).ToString("MM-dd-yyyy");
            var end = new DateTime(year, 12, 31).ToString("MM-dd-yyyy");
            var sql = $@"SELECT 
                                      data_date,
                                      document_type,
                                      success_num,
                                      fail_num
                                    FROM 
                                      public.email_status_summary 
                                    WHERE  data_date >= '{start}' AND data_date <= '{end}'
                                    ORDER BY data_date";

            using (var connection = await Utility.CreateConnection())
            using (var reader = await connection.ExecuteReaderAsync(sql))
            {
                while (reader.Read())
                {
                    var date = reader.GetDateTime(0);
                    var doc = reader.GetString(1);
                    var success = reader.GetInt32(2);
                    var fail = reader.GetInt32(3);
                    result[date] = new Summary { DocType = doc, Success = success, Fail = fail };
                }
            }

            return result;
        }

        private static List<(string branch, string team)> LoadBranch()
        {
            var sep = Path.DirectorySeparatorChar;
            var excel = new FileInfo($"{Environment.CurrentDirectory}{sep}branch.xlsx");
            var result = new List<(string branch, string team)>();
            using (var excelPackage = new ExcelPackage(excel))
            {
                var cells = excelPackage.Workbook.Worksheets[0].Cells;
                for (int i = 3; ; i++)
                {
                    if (cells["A" + i].Value != null)
                    {
                        var b = result.FirstOrDefault(a => a.branch == cells["C" + i].Text);
                        if (b.branch != null)
                        {
                            b.team += "," + cells["C" + i].Text;
                        }
                        else
                        {
                            result.Add((cells["C" + i].Text, cells["A" + i].Text));
                        }

                    }
                    else
                    {
                        break;
                    }
                }
            }

            return result;
        }

       
    }

    public struct Summary
    {
        public string DocType { get; set; }
        public int Success { get; set; }
        public int Fail { get; set; }
    }

    public struct Bounce
    {
        public string Date { get; set; }
        public string DocType { get; set; }
        public string Remark { get; set; }
        public string Account { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string AoName { get; set; }
        public string AoId { get; set; }
        public string Team { get; set; }
    }
}
