using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Hesabdanan.Model;
using Microsoft.AspNetCore.Mvc;


namespace Hesabdanan.Controllers
{
    [Route("/[Controller]")]
    [ApiController]
    public class CreateSellList : Controller
    {
        [HttpPost]
        public async Task<IActionResult> Index(List<IFormFile> files)
        {
            foreach (var file in files)
            {
                if (file == null || file.Length == 0)
                    throw new Exception("FileNotFound");
            }
            List<BankList> bankList = new List<BankList>();
            List<SellerList> sellerList = new List<SellerList>();
            foreach (var excelFile in files)
            {
                var rootFolder = @"C:\Users\zseda\Downloads";
                var fileName = excelFile.FileName;
                var filePath = Path.Combine(rootFolder, fileName);
                var fileLocation = new FileInfo(filePath);

                if (excelFile.Length <= 0)
                    throw new Exception("FileNotFound");


                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var stream = new MemoryStream())
                {
                    excelFile.CopyTo(stream);
                    stream.Position = 0;
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {

                        var rowCount = reader.RowCount;
                        
                        if (fileName == "BankList.xlsx")
                        {
                          
                            for (int i = 1; i <= rowCount; i++)
                            {
                                try
                                {
                                    while (reader.Read()) //Each row of the file
                                    {
                                        bankList.Add(new BankList
                                        {
                                            InvoiceNumber = reader.GetValue(0).ToString(),
                                            Date = reader.GetValue(1).ToString(),
                                            price = reader.GetValue(2).ToString(),
                                            RemainAmount = reader.GetValue(2).ToString(),
                                        });

                                    }
                                }
                                catch (Exception e)
                                {
                                    var hh=i.ToString();
                                    throw;
                                }
                               

                            }
                        }

                        else if (fileName == "SellerList.xlsx")
                        {
                            for (int i = 1; i <= rowCount; i++)
                            {
                                while (reader.Read()) //Each row of the file
                                {
                                    sellerList.Add(new SellerList
                                    {
                                        PId = reader.GetValue(0)?.ToString(),
                                        Date = reader.GetValue(1)?.ToString(),
                                        ProductId = reader.GetValue(2)?.ToString(),
                                        ProductName = reader.GetValue(3)?.ToString(),
                                        Count = reader.GetValue(4)?.ToString(),
                                        UnitPrice = reader.GetValue(5)?.ToString(),
                                        NewCount = reader.GetValue(4)?.ToString(),


                                    });
                                }
                            }
                        }
                    }
                }
            }


            bankList.RemoveAt(0);
            sellerList.RemoveAt(0);
            var resultList = new List<Result>();

            //var unitPriceList = sellerList.Select(w => long.Parse(w.UnitPrice));
            foreach (var item in bankList)
            {
                if (long.Parse(item.InvoiceNumber) == 72)
                {

                }
                var unitPriceList = sellerList.Where(w => long.Parse(w.NewCount) > 0).Select(w => long.Parse(w.UnitPrice));
                var isCountinue = unitPriceList.Where(x => x <= long.Parse(item.RemainAmount)).Any();

                if (isCountinue)
                {
                    for (int i = 0; i < sellerList.Count; i++)
                    {


                        try
                        {
                            if (long.Parse(sellerList[i].NewCount) > 0)
                            {
                                if (isCountinue)
                                {
                                    try
                                    {
                                        var bCount = (long.Parse(item.RemainAmount) / long.Parse(sellerList[i].UnitPrice));
                                        if (bCount < 1)
                                        {

                                            continue;
                                        }
                                        var totalPrice = long.Parse(sellerList[i].Count) * long.Parse(sellerList[i].UnitPrice);
                                        if (bCount <= long.Parse(sellerList[i].NewCount))
                                        {
                                            resultList.Add(new Result
                                            {
                                                InvoiceNumber = item.InvoiceNumber,
                                                PoductId = sellerList[i].ProductId,
                                                ProductName = sellerList[i].ProductName,
                                                SellCount = bCount,
                                                SellUnitPrice = sellerList[i].UnitPrice,
                                                SellAmount = long.Parse(sellerList[i].UnitPrice) * bCount,
                                            });
                                            item.RemainAmount = ((long.Parse(item.RemainAmount)) - long.Parse(sellerList[i].UnitPrice) * bCount).ToString();
                                            sellerList[i].NewCount = (long.Parse(sellerList[i].NewCount) - bCount).ToString();
                                        }
                                        else /*(bCount > long.Parse(sellerList[i].NewCount))*/
                                        {
                                            resultList.Add(new Result
                                            {
                                                InvoiceNumber = item.InvoiceNumber,
                                                PoductId = sellerList[i].ProductId,
                                                ProductName = sellerList[i].ProductName,
                                                SellAmount = (long.Parse(sellerList[i].UnitPrice)) * (long.Parse(sellerList[i].NewCount)),
                                                SellCount = long.Parse(sellerList[i].NewCount),
                                                SellUnitPrice = sellerList[i].UnitPrice,


                                            });
                                            item.RemainAmount = ((long.Parse(item.RemainAmount)) - long.Parse(sellerList[i].UnitPrice) * long.Parse(sellerList[i].NewCount)).ToString();
                                            sellerList[i].NewCount = (long.Parse(sellerList[i].NewCount) - long.Parse(sellerList[i].NewCount)).ToString();
                                        }

                                    }
                                    catch (Exception ex)
                                    {

                                        var hh = "jdj";
                                        throw new Exception(sellerList[i].PId);
                                    }



                                }
                            }
                        }
                        catch (Exception EX)
                        {

                            var hh = "jdj";
                            throw new Exception(sellerList[i].PId);
                        }



                    }


                }
            }



            return Ok();

        }
    }
}

