using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using TechnicalTestMVC.Models;
using LinqToExcel;
using System.Data.OleDb;
using System.Data.Entity.Validation;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;

namespace TechnicalTestMVC.Controllers
{
    public class CustomersController : Controller
    {
        private CustomerDBContext db = new CustomerDBContext();

        // GET: Customers
        public ActionResult Index()
        {
            return View(db.Customers.ToList());
        }

        // GET: Customers/Details/5
        public ActionResult Details(Guid? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Customer customer = db.Customers.Find(id);
            if (customer == null)
            {
                return HttpNotFound();
            }
            return View(customer);
        }

        // GET: Customers/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Customers/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Guid,FullName,DateCreated,Amount,Ref")] Customer customer)
        {
            if (ModelState.IsValid)
            {
                customer.Guid = Guid.NewGuid();
                customer.DateCreated = DateTime.Now;
                db.Customers.Add(customer);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(customer);
        }

        // GET: Customers/Edit/5
        public ActionResult Edit(Guid? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Customer customer = db.Customers.Find(id);
            if (customer == null)
            {
                return HttpNotFound();
            }
            return View(customer);
        }

        // POST: Customers/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Guid,FullName,DateCreated,Amount,Ref")] Customer customer)
        {
            if (ModelState.IsValid)
            {
                db.Entry(customer).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(customer);
        }

        // GET: Customers/Delete/5
        public ActionResult Delete(Guid? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Customer customer = db.Customers.Find(id);
            if (customer == null)
            {
                return HttpNotFound();
            }
            return RedirectToAction("Index");
        }

        // GET: Customers/DeleteALl/5
        public ActionResult DeleteAll(string confirmButton)
        {
                   

            using (var context = new CustomerDBContext())
            {
                var itemsToDelete = context.Set<Customer>();
                context.Customers.RemoveRange(itemsToDelete);
                context.SaveChanges();
            }

            return RedirectToAction("Index");
        }

        // POST: Customers/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(Guid id)
        {
            Customer customer = db.Customers.Find(id);
            db.Customers.Remove(customer);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        // GET: Customers/Import
        public ActionResult Import()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        // Import Excel spreadsheet
        public ActionResult ImportExcel(HttpPostedFileBase FileUpload)
        {
            List<string> data = new List<string>();
            int rowCount = 0;
            if (FileUpload != null)
            {
                // tdata.ExecuteCommand("truncate table OtherCompanyAssets");  
                if (FileUpload.ContentType == "application/vnd.ms-excel" 
                    || FileUpload.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    || FileUpload.ContentType == "application/octet-stream")
                {
                    string filename = FileUpload.FileName;
                    string targetpath = Server.MapPath("~/Uploads/");
                    FileUpload.SaveAs(targetpath + filename);
                    string pathToExcelFile = targetpath + filename;

                    bool firstRow = true;

                    try
                    {
                    //Open the Excel file using ClosedXML.
                    using (XLWorkbook workBook = new XLWorkbook(pathToExcelFile))
                    {
                        //Read the first Sheet from Excel file.
                        IXLWorksheet workSheet = workBook.Worksheet(1);

                        //Create a new DataTable.
                        DataTable dt = new DataTable();
                        //Loop through the Worksheet rows.
                        
                            try
                            {
                                foreach (IXLRow row in workSheet.Rows())
                                {
                                    //Use the first row to add columns to DataTable.
                                    if (firstRow)
                                    {
                                        foreach (IXLCell cell in row.Cells())
                                        {
                                            dt.Columns.Add(cell.Value.ToString());
                                        }
                                        firstRow = false;
                                    }
                                    else
                                    {
                                        //Add rows to DataTable.
                                        dt.Rows.Add();
                                        int i = 0;
                                        foreach (IXLCell cell in row.Cells())
                                        {
                                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                            i++;
                                        }
                                        Guid g = Guid.NewGuid();

                                        string fullName = dt.Rows[dt.Rows.Count - 1][0].ToString();
                                        if (fullName != string.Empty)
                                        {
                                            DateTime dob = Convert.ToDateTime(dt.Rows[dt.Rows.Count - 1][1].ToString());
                                            DateTime dateCreated = Convert.ToDateTime(dt.Rows[dt.Rows.Count - 1][2].ToString());
                                            int reference = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1][3].ToString());
                                            double amount = Convert.ToDouble(dt.Rows[dt.Rows.Count - 1][4].ToString());

                                            Customer cust = new Customer
                                            {
                                                Guid = g,
                                                FullName = fullName,
                                                DateOfBirth = dob,
                                                DateCreated = dateCreated,
                                                Ref = reference,
                                                Amount = amount
                                            };

                                            //add to database
                                            db.Customers.Add(cust);
                                            db.SaveChanges();
                                        }
                                    }
                                    rowCount++;
                                }
                        }
                        catch (DbEntityValidationException ex)
                        {
                            foreach (var entityValidationErrors in ex.EntityValidationErrors)
                            {
                                foreach (var validationError in entityValidationErrors.ValidationErrors)
                                {
                                    Response.Write("Property: " + validationError.PropertyName + " Error: " + validationError.ErrorMessage);
                                }
                            }
                        }
                    }
                        //return Json("success", JsonRequestBehavior.AllowGet);
                        return RedirectToAction("Index");
                    }
                    catch (Exception ex)
                    {
                        string exx = ex.Message + " " + ex.StackTrace;
                        return Json(exx, JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    //alert message for invalid file format  
                    data.Add("<ul>");
                    data.Add("<li>Only Excel file format is allowed</li>");
                    data.Add("</ul>");
                    data.ToArray();
                    return Json(data, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                data.Add("<ul>");
                if (FileUpload == null) data.Add("<li>Please choose Excel file</li>");
                data.Add("</ul>");
                data.ToArray();
                return Json(data, JsonRequestBehavior.AllowGet);
            }
        }
    
      
    }
}
