using Dashboard_New.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using ClosedXML.Excel;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using System.Data;
using Dashboard_New.App_Start;
using System.Web.Security;
using Dashboard_New.Models.custom;

namespace Dashboard_New.Controllers
{

   
    public class HomeController : Controller
    {
        DateTime date1, date2;
        [DashboardAutherisation("chqmdy")]
        //public ActionResult modify_cheques()
        public ActionResult modify_cheques(int page = 1)
        {
            List<chqdrawn> users = new List<chqdrawn>();
            vcEntities mdcc = new vcEntities();
            if (!string.IsNullOrEmpty(Request.Form["from_dt"]) && !string.IsNullOrEmpty(Request.Form["to_dt"]))
            {
                Session["stdt"] = Request.Form["from_dt"];
                date1 = DateTime.ParseExact((string)Session["stdt"], "dd/MM/yyyy", CultureInfo.InvariantCulture);//Convert.ToDateTime();
                Session["todt"] = Request.Form["to_dt"];
                date2 = DateTime.ParseExact((string)Session["todt"], "dd/MM/yyyy", CultureInfo.InvariantCulture);// Convert.ToDateTime(Session["todt"]);
            }
            else
            {
                if (Session["stdt"] != null || Session["todt"] != null)
                {
                    date1 = DateTime.ParseExact((string)Session["stdt"], "dd/MM/yyyy", CultureInfo.InvariantCulture);//Convert.ToDateTime();
                    //Session["todt"] = Request.Form["to_dt"];
                    date2 = DateTime.ParseExact((string)Session["todt"], "dd/MM/yyyy", CultureInfo.InvariantCulture);
                }
                
            }
            chqdrawn model = new chqdrawn();
            model.PageIndex = page;
            model.PageSize = 10;
            model.RecordCount = mdcc.CHECKDRAWNs.Count();
            int startIndex = (page - 1) * model.PageSize;

            using (vcEntities mdc = new vcEntities())
            {

                if (page == 1)
                    users = (from a in mdc.CHECKDRAWNs
                             where a.CQDATE >= date1 && a.CQDATE <= date2
                             select new chqdrawn
                             {
                                 CQDATE = a.CQDATE,
                                 BRNO = a.BRNO,
                                 NPRNO = a.NPRNO,
                                 PARTY = a.PARTY,
                                 CHEQ_NO = a.CHEQ_NO,
                                 RSAMT = a.RSAMT,
                                 VOUCHNO = a.VOUCHNO,
                                 CHEK = a.CHEK
                             }).OrderBy(m => m.CHEQ_NO)
                             .ToList();
                else
                    users = (from a in mdc.CHECKDRAWNs
                             where a.CQDATE >= date1 && a.CQDATE <= date2
                             select new chqdrawn
                             {
                                 CQDATE = a.CQDATE,
                                 BRNO = a.BRNO,
                                 NPRNO = a.NPRNO,
                                 PARTY = a.PARTY,
                                 CHEQ_NO = a.CHEQ_NO,
                                 RSAMT = a.RSAMT,
                                 VOUCHNO = a.VOUCHNO,
                                 CHEK = a.CHEK
                             }).OrderBy(m => m.CHEQ_NO)
                        .Skip(startIndex)
                        //.Take(model.PageSize)
                        .ToList();


            }
            //return View(users);
            return View(new chqmdf() { chqdrawn_dup = users });
        }


        [HttpPost]
        public JsonResult UpdateUser(chqdrawn model)
        {
            string ch = model.CQDATE_dup.ToString();
            ch = ch.Substring(0, 10);

            string to_dt = ch.Trim();
            DateTime todt = DateTime.ParseExact(ch, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            ch = todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);




            //DateTime CQDATE_dupc = Convert.ToDateTime(ch);
            //CQDATE_dupc=CQDATE_dupc.Date;
            //DateTime CQDATE_dupc = DateTime.ParseExact(ch, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            vcEntities mdc = new vcEntities();
            //var upd = (from a in mdc.CHECKDRAWNs
            //            where a.CHEQ_NO == model.CHEQ_NO
            //            select a).FirstOrDefault();
            //var upd = mdc.CHECKDRAWNs.Where(m => string.Equals(m.CHEQ_NO,model.CHEQ_NO_dup.Trim())).SingleOrDefault();
            //var upd = mdc.CHECKDRAWNs.Where(m => string.Equals(m.CHEQ_NO, model.CHEQ_NO_dup.Trim()) && string.Equals(m.VOUCHNO, model.VOUCHNO_dup.Trim())
            //&& string.Equals(m.BRNO, model.BRNO_dup.Trim()) && string.Equals(m.PARTY, model.PARTY_dup.Trim()) && (m.CQDATE == model.CQDATE_dup)
            //&& (m.RSAMT == model.RSAMT_dup) && string.Equals(m.CHEK, model.CHEK_dup.Trim())).SingleOrDefault();
            ////string nari = string.Format("update CHECKDRAWN set  CHEQ_NO = '{0}',VOUCHNO='{1}'  where CHEQ_NO = '{2}'", model.CHEQ_NO, model.VOUCHNO, party_d);
            string query = string.Format("update CHECKDRAWN set  CHEQ_NO = '{0}',VOUCHNO='{1}'  where CHEQ_NO = '{2}' and VOUCHNO='{3}'  and PARTY='{4}'and BRNO='{5}'and CQDATE='{6}'and RSAMT='{7}' AND CHEK='{8}'",
            model.CHEQ_NO, model.VOUCHNO, model.CHEQ_NO_dup, model.VOUCHNO_dup, model.PARTY_dup, model.BRNO_dup, ch, model.RSAMT_dup, model.CHEK_dup);
            mdc.Database.ExecuteSqlCommand(query);
            //upd.CHEQ_NO = model.CHEQ_NO;
            //upd.VOUCHNO = model.VOUCHNO;
            //int count = mdc.SaveChanges();

            //List<chqdrawn> records = new List<chqdrawn>();
            //vcEntities chq_update = new vcEntities();         
            //records = chq_update.Database.SqlQuery<chqdrawn>(string.Format("update CHECKDRAWN set  CHEQ_NO = model.CHEQ_NO,VOUCHNO=a.VOUCHNO  where CQDATE = a.CQDATE and BRNO = a.BRNO,NPRNO = a.NPRNO,PARTY = a.PARTY,CHEQ_NO = a.CHEQ_NO,RSAMT = a.RSAMT,VOUCHNO = a.VOUCHNO,CHEK = a.CHEK")).ToList();
            //Dashboard_New.Models.VModel hg = new VModel();
            //hg.chq = records;
            //return View(new chqmdf() { chqdrawn_dup = records });

            string message = "Success";
            return null;
            //return Json(message, JsonRequestBehavior.AllowGet);
            //return new EmptyResult();
        }



        string myVar;
        // [DashboardAutherisation]
        public ActionResult Index()
        {
            return View();
        }


        [DashboardAutherisation("vc")]
        public ActionResult vc()
        {
            return View(new VModel());
        }

        [HttpPost]
        public ActionResult vcpost(Dashboard_New.Models.custom.vcvendor v1, string grid, string export)
        {
            if (string.IsNullOrEmpty(v1.from_dt))
            {
                ModelState.AddModelError("from_dt", "Date Required");
            }
            if (string.IsNullOrEmpty(v1.to_dt))
            {
                ModelState.AddModelError("to_dt", "Date Required");
            }
            if (ModelState.IsValid)
            {
                if (string.Equals("Export To Excel", export))
                {
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    dt.Columns.AddRange(new DataColumn[18] { new DataColumn("DINP"),new DataColumn("VENDOR NAME"), new DataColumn("VCTRNO"), new DataColumn("AMOUNT"), new DataColumn("NPRNO")
                , new DataColumn("ICCNO"), new DataColumn("COMNO"), new DataColumn("VRNO"), new DataColumn("BRNO"), new DataColumn("VPartyCode"),
                 new DataColumn("ASSTCK"), new DataColumn("ACCTCK"), new DataColumn("ACCT1CK"), new DataColumn("SOCK"), new DataColumn("DRCK"), new DataColumn("CRDATE")
                , new DataColumn("VCTRBNO"), new DataColumn("LUSER")});
                    DateTime fromdt = DateTime.ParseExact(v1.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v1.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.vcvendor>(string.Format("select DINP,VCTRNO,AMOUNT,NPRNO,ICCNO,COMNO,VRNO,BRNO,VPartyCode,ASSTCK,ACCTCK,ACCT1CK,SOCK,DRCK,CRDATE,VCTRBNO,LUSER, convert(decimal,dbo.getAmount(AMOUNT)), (select top 1 VName from vendormaster where VPartyCode = vvv.VPartyCode) VName from vcvouch vvv where DINP>= '{0}' AND DINP<= '{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                    dt.Columns[3].DataType = typeof(decimal);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.DINP.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture), x.vname, x.VCTRNO, x.AMOUNT, x.NPRNO, x.ICCNO, x.COMNO, x.VRNO, x.BRNO, x.VPartyCode, x.ASSTCK, x.ACCTCK, x.ACCT1CK, x.SOCK, x.DRCK, x.CRDATE, x.VCTRBNO, x.LUSER);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("VC");
                        wb.Worksheet(1).Cell(1, 7).Value = "Vendor Credit (VC) : " + v1.from_dt + " to " + v1.to_dt;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "VC.xlsx");
                        }
                    }
                }
                if (string.Equals("Submit", grid))
                {
                    string from_dt = v1.from_dt;
                    string to_dt = v1.to_dt;
                    DateTime fromdt = DateTime.ParseExact(v1.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v1.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    List<vcvendor> records = new List<vcvendor>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<vcvendor>(string.Format("select *, (select top 1 VName from vendormaster where VPartyCode = vvv.VPartyCode) VName from vcvouch vvv where DINP>= '{0}' AND DINP<= '{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel hg = new VModel();
                        hg.vm = records;
                        return View("vc", hg);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Erorrr : " + e); }
                    VModel vd = new VModel();
                    vd.vm = records;
                    return View("vc", vd);
                }
            }
            else
            {
                return View("vc", v1);
            }
            return null;
        }
        [DashboardAutherisation("dc")]
        public ActionResult dc()
        {
            return View(new VModel());
        }
        [HttpPost]
        public ActionResult dcpost(Dashboard_New.Models.custom.dcvendor v2, string grid, string export)
        {
            if (string.IsNullOrEmpty(v2.from_dt1))
            {
                ModelState.AddModelError("from_dt1", "Date Required");
            }
            if (string.IsNullOrEmpty(v2.to_dt1))
            {
                ModelState.AddModelError("to_dt1", "Date Required");
            }
            if (ModelState.IsValid)
            {
                if (string.Equals("Export To Excel", export))
                {
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    dt.Columns.AddRange(new DataColumn[18] { new DataColumn("DINP"),new DataColumn("VENDOR NAME"), new DataColumn("DCTRNO"), new DataColumn("AMOUNT"), new DataColumn("NPRNO")
                , new DataColumn("ICCNO"), new DataColumn("COMNO"), new DataColumn("VRNO"), new DataColumn("BRNO"), new DataColumn("DCID"),
                 new DataColumn("ASSTCK"), new DataColumn("ACCTCK"), new DataColumn("ACCT1CK"), new DataColumn("SOCK"), new DataColumn("DRCK"), new DataColumn("CRDATE")
                , new DataColumn("DCTRBNO"), new DataColumn("LUSER")});
                    DateTime fromdt = DateTime.ParseExact(v2.from_dt1, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v2.to_dt1, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.dcvendor>(string.Format("select DINP,DCTRNO,dbo.getAmount(AMOUNT) AS AMOUNT,NPRNO,ICCNO,COMNO,VRNO,BRNO,DCID,ASSTCK,ACCTCK,ACCT1CK,SOCK,DRCK,CRDATE,DCTRBNO,LUSER, (select top 1 COORNAME from DCMLST where DCID = vvv.DCID) VName from DVOUCH vvv where DINP>= '{0}' AND DINP<= '{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                    dt.Columns[3].DataType = typeof(decimal);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.DINP.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture), x.vname, x.DCTRNO, x.AMOUNT, x.NPRNO, x.ICCNO, x.COMNO, x.VRNO, x.BRNO, x.DCID, x.ASSTCK, x.ACCTCK, x.ACCT1CK, x.SOCK, x.DRCK, x.CRDATE, x.DCTRBNO, x.LUSER);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("DC");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "Direct Credit (DC) : " + v2.from_dt1 + " to " + v2.to_dt1;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "theTitle";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DC.xlsx");
                        }
                    }
                }
                if (string.Equals("Submit", grid))
                {
                    string from_dt = v2.from_dt1;
                    string to_dt = v2.to_dt1;
                    DateTime fromdt = DateTime.ParseExact(v2.from_dt1, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v2.to_dt1, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    List<dcvendor> records = new List<dcvendor>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<dcvendor>(string.Format("select *, (select top 1 COORNAME from DCMLST where DCID = vvv.DCID) VName from DVOUCH vvv where DINP>= '{0}' AND DINP<= '{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.dm = records;
                        return View("dc", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Erorrr : " + e); }
                    VModel vd = new VModel();
                    vd.dm = records;
                    return View("dc", vd);
                }
            }
            else
            {
                return View("dc", v2);
            }
            return null;
        }

        [DashboardAutherisation("nirf")]
        public ActionResult nirf(string id)
        {
            Session["idd"] = id;
            return View(new VModel());
        }

        public ActionResult nirfcons(string id)
        {
            Session["idd"] = id;
            return View(new VModel());
        }
        [HttpPost]
        public ActionResult nirfpost(Dashboard_New.Models.custom.nirf v3, string grid, string export, string id)
        {
            if (string.IsNullOrEmpty(v3.from_year))
            {
                ModelState.AddModelError("From year", "Year Required");
            }
            if (string.IsNullOrEmpty(v3.to_year))
            {
                ModelState.AddModelError("To year", "Year Required");
            }
            if (ModelState.IsValid)
            {
                if ((string.Equals("Export To Excel", export)) && (string.Equals("spons", Session["idd"])))
                {
                    string idd = v3.id;
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    string from_dt = v3.from_year.Substring(2);
                    string to_dt = v3.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v3.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v3.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "REC" + from_dt + to_dt;
                    dt.Columns.AddRange(new DataColumn[9] { new DataColumn("YEAR"), new DataColumn("MONTH"), new DataColumn("PROJECT NUMBER"), new DataColumn("COORDINATOR NAME"), new DataColumn("AGENCY"), new DataColumn("TITLE"), new DataColumn("SANCTION NUMBER"), new DataColumn("SANCTION DATE"), new DataColumn("AMOUNT") });
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.NIRF_RPT>(string.Format("select '" + from_dt + " - " + to_dt + "' as YEAR,DATEPART(MONTH,(R.DATE)) Month,M.NPRNO,M.COOR_NAME,SUBSTRING(M.NPRNO,11,4)AS AGENCY,REPLACE(M.TITLE, CHAR(13) + CHAR(10), '') AS TITLE, M.SANCTNNO, M.SANCTDTE, SUM(R.RT)AS AMOUNT FROM " + tabname + " as  R, MSTLST M WHERE M.NPRNO = R.NPRNO AND((R.RTNO LIKE'P0%') OR(R.RTNO LIKE 'S0%')  OR(R.RTNO LIKE 'M0%')  OR(R.RTNO LIKE 'IH%')OR(R.RTNO LIKE 'NP0%'))AND R.HEAD IS NULL AND SUBSTRING(M.NPRNO, 11, 4)NOT IN(SELECT ResearchCode FROM FOXOFFICE.DBO.InternalProjectCode  WHERE ResearchCode != 'IITM')AND M.NPRNO NOT LIKE'FDR'AND M.NPRNO NOT LIKE'ACC%'AND M.NPRNO NOT LIKE'ICC'AND M.NPRNO NOT LIKE'%DEVP%' AND M.NPRNO NOT LIKE'OAA'AND M.NPRNO NOT LIKE'%EQPT%'AND M.NPRNO NOT LIKE'FDR'AND M.NPRNO NOT LIKE'ICSROH'AND M.NPRNO NOT LIKE'ACC%' AND M.NPRNO NOT LIKE'OTHERS'AND M.NPRNO NOT LIKE'%BMF%'AND M.NPRNO NOT LIKE'%DADM%' AND M.NPRNO NOT LIKE '%DEAN%' AND M.NPRNO NOT LIKE'%DARE%'AND M.NPRNO NOT LIKE'%ACCT%'AND M.NPRNO NOT LIKE'%RMF%' AND M.NPRNO NOT LIKE'%DPLA%' AND M.NPRNO NOT IN('DIA1213001IITMDIAR', 'DIA1718005IITMDIAR', 'DIA1718007ALUMDIAR') AND M.NPRNO NOT LIKE'RSI%' AND M.NPRNO NOT LIKE'%IPRC%' AND M.NPRNO NOT LIKE 'CCE0910008IITMSHAN' AND M.NPRNO NOT LIKE 'COM%' GROUP BY M.NPRNO, M.COOR_NAME, M.SANCTDTE, M.SANCTNNO, DATEPART(MONTH, (R.DATE)), M.TITLE ORDER BY DATEPART(MONTH, (R.DATE)), SUBSTRING(M.NPRNO, 1, 3), M.NPRNO", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();
                    dt.Columns[8].DataType = typeof(Int32);
                    dt.Columns[1].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.YEAR, x.MONTH, x.NPRNO, x.COOR_NAME, x.AGENCY, x.TITLE, x.SANCTNNO, x.SANCTDTE, x.AMOUNT);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("NIRF SPONSORED");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "NIRF(Sponsored) : R" + v3.from_year + " - " + v3.to_year;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "NIRF SPONSORED REPORT";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "NIRF-SPONSORED.xlsx");
                        }
                    }
                }

                if ((string.Equals("Export To Excel", export)) && (string.Equals("nirfcons", Session["idd"])))
                {
                    string idd = v3.id;
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    string from_dt = v3.from_year.Substring(2);
                    string to_dt = v3.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v3.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v3.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "REC" + from_dt + to_dt;
                    dt.Columns.AddRange(new DataColumn[11] { new DataColumn("YEAR"), new DataColumn("DATE"), new DataColumn("MONTH"), new DataColumn("PROJECT NUMBER"), new DataColumn("CONSULTANCY TYPE"), new DataColumn("DEPARTMENT CODE"), new DataColumn("COORDINATOR NAME"), new DataColumn("AGENCY CODE"), new DataColumn("AGENCY"), new DataColumn("TITLE"), new DataColumn("AMOUNT") });
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.NIRF_RPT>(string.Format("select 'REC" + v3.from_year + v3.to_year + "' as YEAR,R.[DATE],DATEPART(MONTH,R.[DATE]) AS MONTH, C.CPRNO,SUBSTRING(C.CPRNO,1,2) AS CONS_TYPE,SUBSTRING(C.CPRNO,7,3) AS DEPT_CODE, C.COOR_NAME1, C.AGENC_CODE AS 'AGENCY_CODE', REPLACE(C.AGENCY, CHAR(13) + CHAR(10), '')  AS AGENCY,   REPLACE(C.C_TITLE, CHAR(13) + CHAR(10), ''), RT FROM " + tabname + " R, CMSTLST C WHERE C.CPRNO = R.ICCNO AND((R.RTNO LIKE'0%') OR(R.RTNO LIKE 'S0%')OR(R.RTNO LIKE 'M0%')) AND C.CPRNO NOT LIKE'IT%'and r.HEAD is null AND C.CPRNO  NOT IN('IC0405001OTERPAYCTEO', 'IC1718001OTERPAYCTEO', 'IC1718002OTERGSTDEAN', 'IC1718ICS003GRANDEAN', 'IC1718IAS004ACCTDEAN') AND R.NPRNO = 'ICC' ORDER BY DATEPART(MONTH, (R.DATE)), SUBSTRING(C.CPRNO, 7, 3)", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();

                    dt.Columns[10].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.YEAR, x.DATE, x.MONTH, x.CPRNO, x.CONS_TYPE, x.DEPT_CODE, x.COOR_NAME, x.AGENCY_CODE, x.AGENCY, x.TITLE, x.rt);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("NIRF-CONSULTANCY");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "NIRF(Consultancy): R" + v3.from_year + " - " + v3.to_year;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "NIRF CONSULTANCY REPORT";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "NIRF-CONSULTANCY.xlsx");
                        }
                    }
                }
                if ((string.Equals("Submit", grid)) && (string.Equals("spons", Session["idd"])))
                {
                    string from_dt = v3.from_year.Substring(2);
                    string to_dt = v3.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v3.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v3.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "REC" + from_dt + to_dt;
                    List<NIRF_RPT> records = new List<NIRF_RPT>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<NIRF_RPT>(string.Format("select '" + v3.from_year + " - " + v3.to_year + "' as YEAR,DATEPART(MONTH,(R.DATE)) Month,M.NPRNO,M.COOR_NAME,SUBSTRING(M.NPRNO,11,4)AS AGENCY,REPLACE(M.TITLE, CHAR(13) + CHAR(10), '') AS TITLE , M.SANCTNNO, M.SANCTDTE, SUM(R.RT)AS AMOUNT FROM " + tabname + " as R, MSTLST M WHERE M.NPRNO = R.NPRNO AND((R.RTNO LIKE'P0%') OR(R.RTNO LIKE 'S0%')  OR(R.RTNO LIKE 'M0%')  OR(R.RTNO LIKE 'IH%')OR(R.RTNO LIKE 'NP0%'))AND R.HEAD IS NULL AND SUBSTRING(M.NPRNO, 11, 4)NOT IN(SELECT ResearchCode FROM FOXOFFICE.DBO.InternalProjectCode  WHERE ResearchCode != 'IITM')AND M.NPRNO NOT LIKE'FDR'AND M.NPRNO NOT LIKE'ACC%'AND M.NPRNO NOT LIKE'ICC'AND M.NPRNO NOT LIKE'%DEVP%' AND M.NPRNO NOT LIKE'OAA'AND M.NPRNO NOT LIKE'%EQPT%'AND M.NPRNO NOT LIKE'FDR'AND M.NPRNO NOT LIKE'ICSROH'AND M.NPRNO NOT LIKE'ACC%' AND M.NPRNO NOT LIKE'OTHERS'AND M.NPRNO NOT LIKE'%BMF%'AND M.NPRNO NOT LIKE'%DADM%' AND M.NPRNO NOT LIKE '%DEAN%' AND M.NPRNO NOT LIKE'%DARE%'AND M.NPRNO NOT LIKE'%ACCT%'AND M.NPRNO NOT LIKE'%RMF%' AND M.NPRNO NOT LIKE'%DPLA%' AND M.NPRNO NOT IN('DIA1213001IITMDIAR', 'DIA1718005IITMDIAR', 'DIA1718007ALUMDIAR') AND M.NPRNO NOT LIKE'RSI%' AND M.NPRNO NOT LIKE'%IPRC%' AND M.NPRNO NOT LIKE 'CCE0910008IITMSHAN' AND M.NPRNO NOT LIKE 'COM%' GROUP BY M.NPRNO, M.COOR_NAME, M.SANCTDTE, M.SANCTNNO, DATEPART(MONTH, (R.DATE)), M.TITLE ORDER BY DATEPART(MONTH, (R.DATE)), SUBSTRING(M.NPRNO, 1, 3), M.NPRNO", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.from_year = v3.from_year;
                        dv.to_year = v3.to_year;
                        dv.nirf = records;
                        return View("nirf", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Error : " + e); }
                    VModel vd = new VModel();
                    vd.nirf = records;
                    vd.from_year = v3.from_year;
                    vd.to_year = v3.to_year;
                    return View("nirf", vd);
                }
                if ((string.Equals("Submit", grid)) && (string.Equals("nirfcons", Session["idd"])))

                {
                    string from_dt = v3.from_year.Substring(2);
                    string to_dt = v3.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v3.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v3.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "REC" + from_dt + to_dt;
                    List<NIRF_RPT> records = new List<NIRF_RPT>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<NIRF_RPT>(string.Format("select 'REC" + v3.from_year + v3.to_year + "' as YEAR,R.[DATE],DATEPART(MONTH,R.[DATE]) AS MONTH, C.CPRNO,SUBSTRING(C.CPRNO,1,2) AS CONS_TYPE,SUBSTRING(C.CPRNO,7,3) AS DEPT_CODE, C.COOR_NAME1, C.AGENC_CODE AS 'AGENCY_CODE', REPLACE(C.AGENCY, CHAR(13) + CHAR(10), '')  AS AGENCY,   REPLACE(C.C_TITLE, CHAR(13) + CHAR(10), '') as TITLE, rt FROM " + tabname + " R, CMSTLST C WHERE C.CPRNO = R.ICCNO AND((R.RTNO LIKE'0%') OR(R.RTNO LIKE 'S0%')OR(R.RTNO LIKE 'M0%')) AND C.CPRNO NOT LIKE'IT%'and r.HEAD is null AND C.CPRNO  NOT IN('IC0405001OTERPAYCTEO', 'IC1718001OTERPAYCTEO', 'IC1718002OTERGSTDEAN', 'IC1718ICS003GRANDEAN', 'IC1718IAS004ACCTDEAN') AND R.NPRNO = 'ICC' ORDER BY DATEPART(MONTH, (R.DATE)), SUBSTRING(C.CPRNO, 7, 3)", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.from_year = v3.from_year;
                        dv.to_year = v3.to_year;
                        dv.nirf = records;
                        return View("nirfcons", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Error : " + e); }
                    VModel vd = new VModel();
                    vd.nirf = records;
                    vd.from_year = v3.from_year;
                    vd.to_year = v3.to_year;
                    return View("nirfcons", vd);
                }
            }
            else
            {
                return View("nirf", v3);
            }
            return null;
        }

        [DashboardAutherisation("ppo")]
        public ActionResult ppo()
        {
            return View(new VModel());
        }
        [HttpPost]
        public ActionResult ppopost(Dashboard_New.Models.custom.ppo v4, string grid, string export)
        {
            if (string.IsNullOrEmpty(v4.from_dt))
            {
                ModelState.AddModelError("From Date", "Year Required");
            }
            if (string.IsNullOrEmpty(v4.to_dt))
            {
                ModelState.AddModelError("To Date", "Year Required");
            }
            if (ModelState.IsValid)
            {
                if (string.Equals("Export To Excel", export))
                {
                    FoxOfficeEntities entities = new FoxOfficeEntities();
                    DataTable dt = new DataTable("Grid");
                    dt.Columns.AddRange(new DataColumn[8] { new DataColumn("INDENT NUMBER"), new DataColumn("INDENT DATE"), new DataColumn("NAME"), new DataColumn("CURRENCY"), new DataColumn("INDENT VALUE"), new DataColumn("DEPARTMENT"), new DataColumn("STATUS"), new DataColumn("REMARKS") });
                    DateTime fromdt = DateTime.ParseExact(v4.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v4.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.ppo>(string.Format("select indent_no,indent_date,[name],currency,indent_value,dept,[status],remarks from indent_master where indent_date>= '{0}' AND indent_date<= '{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                    dt.Columns[4].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.indent_no, x.indent_date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture), x.name, x.currency, x.indent_value, x.dept, x.status, x.remarks);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("Pending_Purchase_Order");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "Pending Indent : " + v4.from_dt + " To " + v4.to_dt;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "Pending Purchase Order Report";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Pending_Indent.xlsx");
                        }
                    }
                }
                if (string.Equals("Submit", grid))
                {
                    //string strddval = Request.Form["menu1"].ToString();
                    string from_dt = v4.from_dt;
                    string to_dt = v4.to_dt;
                    DateTime fromdt = DateTime.ParseExact(v4.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v4.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    List<ppo> records = new List<ppo>();
                    try
                    {
                        FoxOfficeEntities vcobj = new FoxOfficeEntities();
                        records = vcobj.Database.SqlQuery<ppo>(string.Format("select indent_no,indent_date,[name],currency,indent_value,dept,[status],remarks from indent_master where indent_date>= '{0}' AND indent_date<= '{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.ppoo = records;
                        return View("Ppo", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Erorrr : " + e); }
                    VModel vd = new VModel();
                    vd.ppoo = records;
                    return View("Ppo", vd);
                }
            }
            else
            {
                return View("Ppo", v4);
            }
            return null;
        }
        [DashboardAutherisation("pfms_vc")]

        public ActionResult pvc()
        {
            return View(new VModel());
        }

        [HttpPost]
        public ActionResult pvc_post(Dashboard_New.Models.custom.pvc v5, string grid, string export)
        {
            Session["tid"] = Request.Form["Sortby"];
            myVar = Session["tid"].ToString();
            Session["tid1"] = myVar;
            if (string.IsNullOrEmpty(v5.from_year))
            {
                ModelState.AddModelError("From year", "Year Required");
            }
            if (string.IsNullOrEmpty(v5.to_year))
            {
                ModelState.AddModelError("To year", "Year Required");
            }
            if (ModelState.IsValid)
            {
                if ((string.Equals("Export To Excel", export)) && (string.Equals("Year", myVar)))
                {
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    string from_dt = v5.from_year.Substring(2);
                    string to_dt = v5.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v5.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v5.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "VOU" + from_dt + to_dt;
                    dt.Columns.AddRange(new DataColumn[48] { new DataColumn("DATE"), new DataColumn("VRNO"), new DataColumn("NPRNO"), new DataColumn("PART"), new DataColumn("HEAD"), new DataColumn("DISC"), new DataColumn("DIS"), new DataColumn("ICCNO"), new DataColumn("PONO"), new DataColumn("COMNO"), new DataColumn("CQNO"), new DataColumn("BRNO"), new DataColumn("AMOUNT"), new DataColumn("NATURE"), new DataColumn("CHECK"), new DataColumn("REGNO"), new DataColumn("LEDDIS"), new DataColumn("ECODE"), new DataColumn("VCTRNO"), new DataColumn("VPartyCode"), new DataColumn("ASSTCK"), new DataColumn("ACC1TCK"), new DataColumn("ACCTCK"), new DataColumn("SOCK"), new DataColumn("DRCK"), new DataColumn("CRDATE"), new DataColumn("CDSTATUS"), new DataColumn("TRANSFERED"), new DataColumn("EMAILID"), new DataColumn("VCTRBNO"), new DataColumn("LUSER"), new DataColumn("VName"), new DataColumn("VAddress"), new DataColumn("VPinCode"), new DataColumn("VMobile"), new DataColumn("VPhoneNumber"), new DataColumn("VEmailId"), new DataColumn("VPanNo"), new DataColumn("VTinNo"), new DataColumn("VserTaxRegNo"), new DataColumn("VAcctNameInBank"), new DataColumn("VBankName"), new DataColumn("VBranchName"), new DataColumn("VIFSCCode"), new DataColumn("VBankAcctNo"), new DataColumn("VBankMICRCode"), new DataColumn("VbankPhoneNumber"), new DataColumn("VBankEmailID") });
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.pvc>(string.Format("select v.DATE,v.AMOUNT,v.VRNO,v.NPRNO,v.PART,v.HEAD,v.DISC,v.DIS , V.ICCNO, V.PONO, V.COMNO, V.CQNO, V.BRNO, v.NATURE, v.[CHECK], v.REGNO,v.LEDDIS, v.ECODE, c.VCTRNO, c.VPartyCode, c.ASSTCK, c.ACCT1CK, c.ACCTCK,c.SOCK, c.DRCK, c.CRDATE, c.CDSTATUS, c.TRANSFERED, c.EMAILID, c.VCTRBNO, C.LUSER, M.VName,m.VAddress, m.VPinCode, m.VMobile, m.VPhoneNumber, m.VEmailId,m.VPanNo, m.VTinNo, m.VSerTaxRegNo, m.VAcctNameInBank, m.VBankName, m.VBranchName, m.VIFSCCode,M.VBankAcctNo, M.VBankName, M.VBankMICRCode, M.VBankPhoneNumber, M.VBankEmailID  from " + tabname + "  V,VENDORDRAWN C,VendorMaster M WHERE V.VRNO=C.VRNO AND C.VPartyCode=M.VPartyCode and v.nprno in (select nprno from mstlst where ACCOUNTTYPE='pfms')", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();
                    dt.Columns[12].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.DATE, x.VRNO, x.NPRNO, x.PART, x.HEAD, x.DISC, x.DIS, x.ICCNO, x.PONO, x.COMNO, x.CQNO, x.BRNO, x.AMOUNT, x.NATURE, x.CHECK, x.REGNO, x.LEDDIS, x.ECODE, x.VCTRNO, x.VPartyCode, x.ASSTCK, x.ACC1TCK, x.ACCTCK, x.SOCK, x.DRCK, x.CRDATE, x.CDSTATUS, x.TRANSFERED, x.EMAILID, x.VCTRBNO, x.LUSER, x.VName, x.VAddress, x.VPinCode, x.VMobile, x.VPhoneNumber, x.VEmailId, x.VPanNo, x.VTinNo, x.VserTaxRegNo, x.VAcctNameInBank, x.VBankName, x.VBranchName, x.VIFSCCode, x.VBankAcctNo, x.VBankMICRCode, x.VbankPhoneNumber, x.VBankEmailID);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("PFMS-VC-Yearwise");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "PFMS (VC-Yearwise) : VOU" + v5.from_year + " - " + v5.to_year;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "PFMS VC REPORT- Yearwise";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PFMS-VC-Yearwise.xlsx");
                        }
                    }
                }
                if ((string.Equals("Export To Excel", export)) && (string.Equals("Date", myVar)))
                {
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    string from_dt = v5.from_year.Substring(8);
                    string to_dt = v5.to_year.Substring(3, 2);
                    int from_NT = Int32.Parse(from_dt) + 1;
                    DateTime fromdt = DateTime.ParseExact(v5.from_year.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v5.to_year.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string fm, tm;
                    fm = fromdt.ToString("yyyy-MM-dd");
                    tm = todt.ToString("yyyy-MM-dd");
                    string tabname = "VOU" + from_dt + from_NT.ToString();
                    dt.Columns.AddRange(new DataColumn[48] { new DataColumn("DATE"), new DataColumn("VRNO"), new DataColumn("NPRNO"), new DataColumn("PART"), new DataColumn("HEAD"), new DataColumn("DISC"), new DataColumn("DIS"), new DataColumn("ICCNO"), new DataColumn("PONO"), new DataColumn("COMNO"), new DataColumn("CQNO"), new DataColumn("BRNO"), new DataColumn("AMOUNT"), new DataColumn("NATURE"), new DataColumn("CHECK"), new DataColumn("REGNO"), new DataColumn("LEDDIS"), new DataColumn("ECODE"), new DataColumn("VCTRNO"), new DataColumn("VPartyCode"), new DataColumn("ASSTCK"), new DataColumn("ACC1TCK"), new DataColumn("ACCTCK"), new DataColumn("SOCK"), new DataColumn("DRCK"), new DataColumn("CRDATE"), new DataColumn("CDSTATUS"), new DataColumn("TRANSFERED"), new DataColumn("EMAILID"), new DataColumn("VCTRBNO"), new DataColumn("LUSER"), new DataColumn("VName"), new DataColumn("VAddress"), new DataColumn("VPinCode"), new DataColumn("VMobile"), new DataColumn("VPhoneNumber"), new DataColumn("VEmailId"), new DataColumn("VPanNo"), new DataColumn("VTinNo"), new DataColumn("VserTaxRegNo"), new DataColumn("VAcctNameInBank"), new DataColumn("VBankName"), new DataColumn("VBranchName"), new DataColumn("VIFSCCode"), new DataColumn("VBankAcctNo"), new DataColumn("VBankMICRCode"), new DataColumn("VbankPhoneNumber"), new DataColumn("VBankEmailID") });
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.pvc>(string.Format("select v.DATE,v.AMOUNT,v.VRNO,v.NPRNO,v.PART,v.HEAD,v.DISC,v.DIS, V.ICCNO, V.PONO, V.COMNO, V.CQNO, V.BRNO, v.NATURE, v.[CHECK], v.REGNO,v.LEDDIS, v.ECODE, c.VCTRNO, c.VPartyCode, c.ASSTCK, c.ACCT1CK, c.ACCTCK,c.SOCK, c.DRCK, c.CRDATE, c.CDSTATUS, c.TRANSFERED, c.EMAILID, c.VCTRBNO, C.LUSER, M.VName,m.VAddress, m.VPinCode, m.VMobile, m.VPhoneNumber, m.VEmailId,m.VPanNo, m.VTinNo, m.VSerTaxRegNo, m.VAcctNameInBank, m.VBankName, m.VBranchName, m.VIFSCCode,M.VBankAcctNo, M.VBankName, M.VBankMICRCode, M.VBankPhoneNumber, M.VBankEmailID from " + tabname + " V, VENDORDRAWN C, VendorMaster M WHERE V.VRNO = C.VRNO AND C.VPartyCode = M.VPartyCode and v.nprno in (select nprno from mstlst where ACCOUNTTYPE = 'pfms')  and v.DATE>='{0}' AND v.DATE<= '{1}'", fm, tm)).ToList();
                    dt.Columns[12].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.DATE, x.VRNO, x.NPRNO, x.PART, x.HEAD, x.DISC, x.DIS, x.ICCNO, x.PONO, x.COMNO, x.CQNO, x.BRNO, x.AMOUNT, x.NATURE, x.CHECK, x.REGNO, x.LEDDIS, x.ECODE, x.VCTRNO, x.VPartyCode, x.ASSTCK, x.ACC1TCK, x.ACCTCK, x.SOCK, x.DRCK, x.CRDATE, x.CDSTATUS, x.TRANSFERED, x.EMAILID, x.VCTRBNO, x.LUSER, x.VName, x.VAddress, x.VPinCode, x.VMobile, x.VPhoneNumber, x.VEmailId, x.VPanNo, x.VTinNo, x.VserTaxRegNo, x.VAcctNameInBank, x.VBankName, x.VBranchName, x.VIFSCCode, x.VBankAcctNo, x.VBankMICRCode, x.VbankPhoneNumber, x.VBankEmailID);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("PFMS -VC - Datewise");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "PFMS (VC-Datewise) : " + tabname;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "PFMS VC REPORT-Datewise";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PFMS-VC-datewise.xlsx");
                        }
                    }
                }
                if ((string.Equals("Submit", grid)) && (string.Equals("Year", myVar)))
                {
                    string from_yr = v5.from_year.Substring(2);
                    string to_yr = v5.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v5.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v5.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "VOU" + from_yr + to_yr;
                    List<pvc> records = new List<pvc>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<pvc>(string.Format("select v.DATE,v.AMOUNT,v.VRNO,v.NPRNO,v.PART,v.HEAD,v.DISC,v.DIS , V.ICCNO, V.PONO, V.COMNO, V.CQNO, V.BRNO, v.NATURE, v.[CHECK], v.REGNO,v.LEDDIS, v.ECODE, c.VCTRNO, c.VPartyCode, c.ASSTCK, c.ACCT1CK, c.ACCTCK,c.SOCK, c.DRCK, c.CRDATE, c.CDSTATUS, c.TRANSFERED, c.EMAILID, c.VCTRBNO, C.LUSER, M.VName,m.VAddress, m.VPinCode, m.VMobile, m.VPhoneNumber, m.VEmailId,m.VPanNo, m.VTinNo, m.VSerTaxRegNo, m.VAcctNameInBank, m.VBankName, m.VBranchName, m.VIFSCCode,M.VBankAcctNo, M.VBankName, M.VBankMICRCode, M.VBankPhoneNumber, M.VBankEmailID  from " + tabname + "  V,VENDORDRAWN C,VendorMaster M WHERE V.VRNO=C.VRNO AND C.VPartyCode=M.VPartyCode and v.nprno in (select nprno from mstlst where ACCOUNTTYPE='pfms')", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.from_year = v5.from_year;
                        dv.to_year = v5.to_year;
                        dv.pfm_vclist = records;
                        return View("pvc", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Error : " + e); }
                    VModel vd = new VModel();
                    vd.pfm_vclist = records;
                    vd.from_year = v5.from_year;
                    vd.to_year = v5.to_year;
                    return View("pvc", vd);
                }
                if ((string.Equals("Submit", grid)) && (string.Equals("Date", myVar)))
                {
                    string from_dt = v5.from_year.Substring(8);
                    int from_NT = Int32.Parse(from_dt) + 1;
                    string to_dt = v5.to_year.Substring(3, 2);
                    DateTime fromdt = DateTime.ParseExact(v5.from_year.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v5.to_year.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string fm, tm;
                    fm = fromdt.ToString("yyyy-MM-dd");
                    tm = todt.ToString("yyyy-MM-dd");
                    string tabname = "VOU" + from_dt + from_NT.ToString();
                    List<pvc> records = new List<pvc>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<pvc>(string.Format("select v.DATE,v.AMOUNT,v.VRNO,v.NPRNO,v.PART,v.HEAD,v.DISC,v.DIS, V.ICCNO, V.PONO, V.COMNO, V.CQNO, V.BRNO, v.NATURE, v.[CHECK], v.REGNO,v.LEDDIS, v.ECODE, c.VCTRNO, c.VPartyCode, c.ASSTCK, c.ACCT1CK, c.ACCTCK,c.SOCK, c.DRCK, c.CRDATE, c.CDSTATUS, c.TRANSFERED, c.EMAILID, c.VCTRBNO, C.LUSER, M.VName,m.VAddress, m.VPinCode, m.VMobile, m.VPhoneNumber, m.VEmailId,m.VPanNo, m.VTinNo, m.VSerTaxRegNo, m.VAcctNameInBank, m.VBankName, m.VBranchName, m.VIFSCCode,M.VBankAcctNo, M.VBankName, M.VBankMICRCode, M.VBankPhoneNumber, M.VBankEmailID from " + tabname + " V, VENDORDRAWN C, VendorMaster M WHERE V.VRNO = C.VRNO AND C.VPartyCode = M.VPartyCode and v.nprno in (select nprno from mstlst where ACCOUNTTYPE = 'pfms')  and v.DATE between '{0}' AND '{1}'", fm, tm)).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.from_year = v5.from_year;
                        dv.to_year = v5.to_year;
                        dv.pfm_vclist = records;
                        return View("pvc", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Error : " + e); }
                    VModel vd = new VModel();
                    vd.pfm_vclist = records;
                    vd.from_year = v5.from_year;
                    vd.to_year = v5.to_year;
                    return View("pvc", vd);
                }
            }
            else
            {
                return View("pvc", v5);
            }
            return null;
        }
        [DashboardAutherisation("pfms_dc")]

        public ActionResult pdc()
        {
            return View(new VModel());
        }
        [HttpPost]
        public ActionResult pdc_post(Dashboard_New.Models.custom.pdc v6, string grid, string export)
        {
            Session["tid"] = Request.Form["Sortby"];
            myVar = Session["tid"].ToString();
            Session["tid1"] = myVar;
            if (string.IsNullOrEmpty(v6.from_year))
            {
                ModelState.AddModelError("From year", "Year Required");
            }
            if (string.IsNullOrEmpty(v6.to_year))
            {
                ModelState.AddModelError("To year", "Year Required");
            }
            if (ModelState.IsValid)
            {
                if ((string.Equals("Export To Excel", export)) && (string.Equals("Year", myVar)))
                {
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    string from_dt = v6.from_year.Substring(2);
                    string to_dt = v6.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v6.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v6.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "VOU" + from_dt + to_dt;
                    dt.Columns.AddRange(new DataColumn[40] { new DataColumn("DATE"), new DataColumn("VRNO"), new DataColumn("NPRNO"), new DataColumn("PART"), new DataColumn("HEAD"), new DataColumn("DISC"), new DataColumn("DIS"), new DataColumn("ICCNO"), new DataColumn("PONO"), new DataColumn("COMNO"), new DataColumn("CQNO"), new DataColumn("OPTION"), new DataColumn("BRNO"), new DataColumn("AMOUNT"), new DataColumn("NATURE"), new DataColumn("CHECK"), new DataColumn("REGNO"), new DataColumn("LEDDIS"), new DataColumn("ECODE"), new DataColumn("DCTRNO"), new DataColumn("DCID"), new DataColumn("ASSTCK"), new DataColumn("ACC1TCK"), new DataColumn("ACCTCK"), new DataColumn("SOCK"), new DataColumn("DRCK"), new DataColumn("CRDATE"), new DataColumn("CDSTATUS"), new DataColumn("TRANSFERED"), new DataColumn("EMAILID"), new DataColumn("DCTRBNO"), new DataColumn("INSTID"), new DataColumn("COORNAME"), new DataColumn("DEPTNAME"), new DataColumn("BANKTYPE"), new DataColumn("VPhoneNumber"), new DataColumn("CBANKACCTNO"), new DataColumn("COOREMAILID"), new DataColumn("LUSER"), new DataColumn("ACCOUNTTYPE") });

                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.pdc>(string.Format("select * from  " + tabname + " v,CREDITDRAWN c,DCMLST d WHERE v.vrno=c.vrno and c.DCID=d.DCID and v.nprno in (select nprno from mstlst where ACCOUNTTYPE='pfms')", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();
                    dt.Columns[13].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.DATE, x.VRNO, x.NPRNO, x.PART, x.HEAD, x.DISC, x.DIS, x.ICCNO, x.PONO, x.COMNO, x.CQNO, x.OPTION, x.BRNO, x.AMOUNT, x.NATURE, x.CHECK, x.REGNO, x.LEDDIS, x.ECODE, x.DCTRNO, x.ASSTCK, x.ACC1TCK, x.ACCTCK, x.SOCK, x.DRCK, x.CRDATE, x.CDSTATUS, x.TRANSFERED, x.EMAILID, x.DCTRBNO, x.INSTID, x.COORNAME, x.DEPTNAME, x.BANKTYPE, x.VPhoneNumber, x.CBANKACCTNO, x.COOREMAILID, x.LUSER, x.ACCOUNTTYPE);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("PFMS-DC-Yearwise");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "PFMS (DC-Yearwise) : VOU" + v6.from_year + " - " + v6.to_year;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "PFMS DC REPORT- Yearwise";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PFMS-DC-Yearwise.xlsx");
                        }
                    }
                }
                if ((string.Equals("Export To Excel", export)) && (string.Equals("Date", myVar)))
                {
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    string from_dt = v6.from_year.Substring(8);
                    string to_dt = v6.to_year.Substring(3, 2);
                    int from_NT = Int32.Parse(from_dt) + 1;
                    DateTime fromdt = DateTime.ParseExact(v6.from_year.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v6.to_year.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string fm, tm;
                    fm = fromdt.ToString("yyyy-MM-dd");
                    tm = todt.ToString("yyyy-MM-dd");
                    string tabname = "VOU" + from_dt + from_NT.ToString();
                    dt.Columns.AddRange(new DataColumn[40] { new DataColumn("DATE"), new DataColumn("VRNO"), new DataColumn("NPRNO"), new DataColumn("PART"), new DataColumn("HEAD"), new DataColumn("DISC"), new DataColumn("DIS"), new DataColumn("ICCNO"), new DataColumn("PONO"), new DataColumn("COMNO"), new DataColumn("CQNO"), new DataColumn("OPTION"), new DataColumn("BRNO"), new DataColumn("AMOUNT"), new DataColumn("NATURE"), new DataColumn("CHECK"), new DataColumn("REGNO"), new DataColumn("LEDDIS"), new DataColumn("ECODE"), new DataColumn("DCTRNO"), new DataColumn("DCID"), new DataColumn("ASSTCK"), new DataColumn("ACC1TCK"), new DataColumn("ACCTCK"), new DataColumn("SOCK"), new DataColumn("DRCK"), new DataColumn("CRDATE"), new DataColumn("CDSTATUS"), new DataColumn("TRANSFERED"), new DataColumn("EMAILID"), new DataColumn("DCTRBNO"), new DataColumn("INSTID"), new DataColumn("COORNAME"), new DataColumn("DEPTNAME"), new DataColumn("BANKTYPE"), new DataColumn("VPhoneNumber"), new DataColumn("CBANKACCTNO"), new DataColumn("COOREMAILID"), new DataColumn("LUSER"), new DataColumn("ACCOUNTTYPE") });
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.pdc>(string.Format("select * from  " + tabname + " v,CREDITDRAWN c,DCMLST d WHERE v.vrno=c.vrno and c.DCID=d.DCID and v.nprno in (select nprno from mstlst where ACCOUNTTYPE='pfms')  and v.DATE>='{0}' AND v.DATE<= '{1}'", fm, tm)).ToList();
                    dt.Columns[13].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.DATE, x.VRNO, x.NPRNO, x.PART, x.HEAD, x.DISC, x.DIS, x.ICCNO, x.PONO, x.COMNO, x.CQNO, x.OPTION, x.BRNO, x.AMOUNT, x.NATURE, x.CHECK, x.REGNO, x.LEDDIS, x.ECODE, x.DCTRNO, x.ASSTCK, x.ACC1TCK, x.ACCTCK, x.SOCK, x.DRCK, x.CRDATE, x.CDSTATUS, x.TRANSFERED, x.EMAILID, x.DCTRBNO, x.INSTID, x.COORNAME, x.DEPTNAME, x.BANKTYPE, x.VPhoneNumber, x.CBANKACCTNO, x.COOREMAILID, x.LUSER, x.ACCOUNTTYPE);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("PFMS -DC - Datewise");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "PFMS (DC-Datewise) : " + tabname;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "PFMS DC REPORT-Datewise";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PFMS-DC-datewise.xlsx");
                        }
                    }
                }
                if ((string.Equals("Submit", grid)) && (string.Equals("Year", myVar)))
                {
                    string from_yr = v6.from_year.Substring(2);
                    string to_yr = v6.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v6.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v6.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "VOU" + from_yr + to_yr;
                    List<pdc> records = new List<pdc>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<pdc>(string.Format("select * from  " + tabname + " v,CREDITDRAWN c,DCMLST d WHERE v.vrno=c.vrno and c.DCID=d.DCID and v.nprno in (select nprno from mstlst where ACCOUNTTYPE='pfms')", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.from_year = v6.from_year;
                        dv.to_year = v6.to_year;
                        dv.pfm_dclist = records;
                        return View("pdc", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Error : " + e); }
                    VModel vd = new VModel();
                    vd.pfm_dclist = records;
                    vd.from_year = v6.from_year;
                    vd.to_year = v6.to_year;
                    return View("pdc", vd);
                }
                if ((string.Equals("Submit", grid)) && (string.Equals("Date", myVar)))
                {
                    string from_dt = v6.from_year.Substring(8);
                    int from_NT = Int32.Parse(from_dt) + 1;
                    string to_dt = v6.to_year.Substring(3, 2);
                    DateTime fromdt = DateTime.ParseExact(v6.from_year.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v6.to_year.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string fm, tm;
                    fm = fromdt.ToString("yyyy-MM-dd");
                    tm = todt.ToString("yyyy-MM-dd");
                    string tabname = "VOU" + from_dt + from_NT.ToString();
                    List<pdc> records = new List<pdc>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<pdc>(string.Format("select * from  " + tabname + " v,CREDITDRAWN c,DCMLST d WHERE v.vrno=c.vrno and c.DCID=d.DCID and v.nprno in (select nprno from mstlst where ACCOUNTTYPE='pfms') and v.DATE between '{0}' AND '{1}'", fm, tm)).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.from_year = v6.from_year;
                        dv.to_year = v6.to_year;
                        dv.pfm_dclist = records;
                        return View("pdc", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Error : " + e); }
                    VModel vd = new VModel();
                    vd.pfm_dclist = records;
                    vd.from_year = v6.from_year;
                    vd.to_year = v6.to_year;
                    return View("pdc", vd);
                }
            }
            else
            {
                return View("pdc", v6);
            }
            return null;
        }
        [DashboardAutherisation("bill")]

        public ActionResult bill(string id)
        {
            Session["bid"] = id;
            return View(new VModel());
        }

        public ActionResult billcons(string id)
        {
            Session["bid"] = id;
            return View(new VModel());
        }

        [HttpPost]
        public ActionResult billspons_post(Dashboard_New.Models.custom.bill v7, string grid, string export, string id)
        {
            if (string.IsNullOrEmpty(v7.from_dt))
            {
                ModelState.AddModelError("From date", "Year Required");
            }
            if (string.IsNullOrEmpty(v7.to_dt))
            {
                ModelState.AddModelError("To date", "Year Required");
            }
            if (ModelState.IsValid)
            {
                if ((string.Equals("Export To Excel", export)) && (string.Equals("spons", Session["bid"])))
                {
                    string idd = v7.id;
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    DateTime fromdt = DateTime.ParseExact(v7.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v7.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string fm, tm;
                    fm = fromdt.ToString("yyyy-MM-dd");
                    tm = todt.ToString("yyyy-MM-dd");
                    dt.Columns.AddRange(new DataColumn[46] { new DataColumn("BILL DATE"), new DataColumn("BILL NO"), new DataColumn("CASHDATE"), new DataColumn("BILLNO_DT"), new DataColumn("NPRNO"), new DataColumn("HEAD"), new DataColumn("COMNO"), new DataColumn("PONO"), new DataColumn("VOCHNO"), new DataColumn("VAMOUNT"), new DataColumn("PARTY"), new DataColumn("ITEM"), new DataColumn("ADD1"), new DataColumn("ADD2"), new DataColumn("ADD3"), new DataColumn("CITY"), new DataColumn("NARATION1"), new DataColumn("NARATION2"), new DataColumn("SPNSTOCSH"), new DataColumn("SPNSTOCSH1"), new DataColumn("CQCK"), new DataColumn("LEDGER1"), new DataColumn("LEDGER"), new DataColumn("LEDGER2"), new DataColumn("SCHOLAR"), new DataColumn("TIME"), new DataColumn("ADJ"), new DataColumn("ECODE"), new DataColumn("IT"), new DataColumn("PT"), new DataColumn("PARTYCODE"), new DataColumn("PANNUMBER"), new DataColumn("TPNO"), new DataColumn("TPLDATE"), new DataColumn("ACCOUNTTYPE"), new DataColumn("TDSPERSENT"), new DataColumn("TDSSECTION"), new DataColumn("TDSAMOUNT"), new DataColumn("PARTYAMOUNT"), new DataColumn("TAXABLEAMOUNT"), new DataColumn("INVOICEAMOUNT"), new DataColumn("VOUCHERNUMBER"), new DataColumn("VOUCHERDATE"), new DataColumn("TDSGSTVALUE"), new DataColumn("TDSGSTAMOUNT"), new DataColumn("SLNO") });
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.bill>(string.Format("select BRDATE,BRNO,CASHDATE,NPRNO,HEAD,COMNO,PONO,VOCHNO,VAMOUNT,ITEM,PARTY,ADD1, ADD2,ADD3,CITY,NARATION1,NARATION2,SPNSTOCSH,SPNSTOCSH1,CQCK,LEDGER1,LEDGER,LEDGER2,SCHOLAR,TIME,ADJ,ECODE,IT,PT,PARTYCODE,PANNUMBER,TPNO,TPLDATE,ACCOUNTTYPE,TDSPERSENT,TDSSECTION,TDSAMOUNT,PARTYAMOUNT,TAXABLEAMOUNT,INVOICEAMOUNT,VOUCHERNUMBER, VOUCHERDATE,TDSGSTVALUE,TDSGSTAMOUNT,SLNO from br where  brdate >='{0}' and brdate <='{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                    dt.Columns[9].DataType = typeof(Int32);
                    dt.Columns[38].DataType = typeof(Int32);
                    dt.Columns[39].DataType = typeof(Int32);
                    dt.Columns[40].DataType = typeof(Int32);
                    dt.Columns[43].DataType = typeof(Int32);
                    dt.Columns[44].DataType = typeof(Int32);
                    dt.Columns[45].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.BRDATE, x.BRNO, x.CASHDATE, x.BILLNO_DT, x.NPRNO, x.HEAD, x.COMNO, x.PONO, x.VOCHNO, x.VAMOUNT, x.ITEM, x.PARTY, x.ADD1, x.ADD2, x.ADD3, x.CITY, x.NARATION1, x.NARATION2, x.SPNSTOCSH, x.SPNSTOCSH1, x.CQCK, x.LEDGER1, x.LEDGER, x.LEDGER2, x.SCHOLAR, x.TIME, x.ADD1, x.ECODE, x.IT, x.PT, x.PARTYCODE, x.PANNUMBER, x.TPNO, x.TPLDATE, x.ACCOUNTTYPE, x.TDSPERSENT, x.TDSSECTION, x.TDSAMOUNT, x.PARTYAMOUNT, (x.TAXABLEAMOUNT), x.INVOICEAMOUNT, x.VOUCHERNUMBER, x.VOUCHERDATE, x.TDSGSTVALUE, x.TDSGSTAMOUNT, x.SLNO);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("BILL SPONSORED");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "BILL(Sponsored) : " + v7.from_dt + " - " + v7.to_dt;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "BILL SPONSORED REPORT";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "BILL-SPONSORED.xlsx");
                        }
                    }
                }
                if ((string.Equals("Export To Excel", export)) && (string.Equals("billcons", Session["bid"])))
                {
                    string idd = v7.id;
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    DateTime fromdt = DateTime.ParseExact(v7.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v7.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string fm, tm;
                    fm = fromdt.ToString("yyyy-MM-dd");
                    tm = todt.ToString("yyyy-MM-dd");
                    dt.Columns.AddRange(new DataColumn[40] { new DataColumn("BILL DATE"), new DataColumn("BILL NO"), new DataColumn("CASHDATE"), new DataColumn("NPRNO"), new DataColumn("HEAD"), new DataColumn("COMNO"), new DataColumn("PONO"), new DataColumn("VOCHNO"), new DataColumn("VAMOUNT"), new DataColumn("PARTY"), new DataColumn("ITEM"), new DataColumn("ADD1"), new DataColumn("CITY"), new DataColumn("NARATION1"), new DataColumn("NARATION2"), new DataColumn("SPNSTOCSH"), new DataColumn("SPNSTOCSH1"), new DataColumn("SPNSTOCSH2"), new DataColumn("CQCK"), new DataColumn("LEDGER"), new DataColumn("SCHOLAR"), new DataColumn("TIME"), new DataColumn("ADJ"), new DataColumn("ECODE"), new DataColumn("PARTYCODE"), new DataColumn("PANNUMBER"), new DataColumn("TPNO"), new DataColumn("TPLDATE"), new DataColumn("ACCOUNTTYPE"), new DataColumn("TDSPERSENT"), new DataColumn("TDSSECTION"), new DataColumn("TDSAMOUNT"), new DataColumn("PARTYAMOUNT"), new DataColumn("TAXABLEAMOUNT"), new DataColumn("INVOICEAMOUNT"), new DataColumn("VOUCHERNUMBER"), new DataColumn("VOUCHERDATE"), new DataColumn("TDSGSTVALUE"), new DataColumn("TDSGSTAMOUNT"), new DataColumn("SLNO") });
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.bill>(string.Format("select BRDATE,BRNO,CASHDATE,BILLNO_DT,NPRNO,HEAD,COMNO,PONO,VOCHNO,VAMOUNT,ITEM,PARTY,ADD1,CITY,NARATION1,NARATION2,SPNSTOCSH1,SPNSTOCSH,SPNSTOCSH2,CQCK,LEDGER,TIME,ADJ,ECODE,PARTYCODE,PANNUMBER,TPLNO,TPLDATE,ACCOUNTTYPE,TDSPERSENT,TDSSECTION,TDSAMOUNT,PARTYAMOUNT,TAXABLEAMOUNT,INVOICEAMOUNT,VOUCHERNUMBER,VOUCHERDATE,TDSGSTVALUE,TDSGSTAMOUNT,SLNO from CONbr where  brdate >='{0}' and brdate <='{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                    dt.Columns[8].DataType = typeof(Int32);
                    dt.Columns[32].DataType = typeof(Int32);
                    dt.Columns[33].DataType = typeof(Int32);
                    dt.Columns[34].DataType = typeof(Int32);
                    dt.Columns[35].DataType = typeof(Int32);
                    dt.Columns[38].DataType = typeof(Int32);
                    // dt.Columns[39].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.BILLNO_DT, x.BRNO, x.CASHDATE, x.NPRNO, x.HEAD, x.COMNO, x.PONO, x.VOCHNO, x.VAMOUNT, x.PARTY, x.ITEM, x.ADD1, x.CITY, x.NARATION1, x.NARATION2, x.SPNSTOCSH, x.SPNSTOCSH1, x.SPNSTOCSH2, x.CQCK, x.LEDGER, x.SCHOLAR, x.TIME, x.ADJ, x.ECODE, x.PARTYCODE, x.PANNUMBER, x.TPNO, x.TPLDATE, x.ACCOUNTTYPE, x.TDSPERSENT, x.TDSSECTION, x.TDSAMOUNT, x.PARTYAMOUNT, x.TAXABLEAMOUNT, x.INVOICEAMOUNT, x.VOUCHERNUMBER, x.VOUCHERDATE, x.TDSGSTVALUE, x.TDSGSTAMOUNT, x.SLNO);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("CONSULTANCY BILL SPONSORED");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "CONSULTANCY BILL(Sponsored) : " + v7.from_dt + " - " + v7.to_dt;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "CONSULTANCY BILL SPONSORED REPORT";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CONSULTANCY BILL-SPONSORED.xlsx");
                        }
                    }
                }
                if ((string.Equals("Submit", grid)) && (string.Equals("spons", Session["bid"])))
                {
                    string from_dt = v7.from_dt;
                    string to_dt = v7.to_dt;
                    DateTime fromdt = DateTime.ParseExact(v7.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v7.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string fm, tm;
                    fm = fromdt.ToString("yyyy-MM-dd");
                    tm = todt.ToString("yyyy-MM-dd");
                    List<bill> records = new List<bill>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<bill>(string.Format("select BRDATE,BRNO,CASHDATE,BILLNO_DT,NPRNO,HEAD,COMNO,PONO,VOCHNO,VAMOUNT,ITEM,PARTY,ADD1, ADD2,ADD3,CITY,NARATION1,NARATION2,SPNSTOCSH,SPNSTOCSH1,CQCK,LEDGER1,LEDGER,LEDGER2,SCHOLAR,TIME,ADJ,ECODE,IT,PT,PARTYCODE,PANNUMBER,TPNO,TPLDATE,ACCOUNTTYPE,TDSPERSENT,TDSSECTION,TDSAMOUNT,PARTYAMOUNT,TAXABLEAMOUNT,INVOICEAMOUNT,VOUCHERNUMBER, VOUCHERDATE,TDSGSTVALUE,TDSGSTAMOUNT,SLNO from br where  brdate >='{0}' and brdate <='{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.from_year = v7.from_dt;
                        dv.to_year = v7.to_dt;
                        dv.bill_sp = records;
                        return View("bill", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Error : " + e); }
                    VModel vd = new VModel();
                    vd.bill_sp = records;
                    vd.from_year = v7.from_year;
                    vd.to_year = v7.to_year;
                    return View("bill", vd);
                }
                if ((string.Equals("Submit", grid)) && (string.Equals("billcons", Session["bid"])))
                {
                    string from_dt = v7.from_dt;
                    string to_dt = v7.to_dt;
                    DateTime fromdt = DateTime.ParseExact(v7.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v7.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    string fm, tm;
                    fm = fromdt.ToString("yyyy-MM-dd");
                    tm = todt.ToString("yyyy-MM-dd");
                    List<billcons> records = new List<billcons>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<billcons>(string.Format("select BRDATE,BRNO,CASHDATE,BILLNO_DT,NPRNO,HEAD,COMNO,PONO,VOCHNO,VAMOUNT,ITEM,PARTY,ADD1,CITY,NARATION1,NARATION2,SPNSTOCSH1,SPNSTOCSH,SPNSTOCSH2,CQCK,LEDGER,TIME,ADJ,ECODE,PARTYCODE,PANNUMBER,TPLNO,TPLDATE,ACCOUNTTYPE,TDSPERSENT,TDSSECTION,TDSAMOUNT,PARTYAMOUNT,TAXABLEAMOUNT,INVOICEAMOUNT,VOUCHERNUMBER,VOUCHERDATE,TDSGSTVALUE,TDSGSTAMOUNT,SLNO from conbr where  brdate >='{0}' and brdate <='{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.from_year = v7.from_dt;
                        dv.to_year = v7.to_dt;
                        dv.bill_con = records;
                        return View("billcons", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Error : " + e); }
                    VModel vd = new VModel();
                    vd.bill_con = records;
                    vd.from_year = v7.from_year;
                    vd.to_year = v7.to_year;
                    return View("billcons", vd);
                }
            }
            else
            {
                return View("bill", v7);
            }
            return null;
        }

        [DashboardAutherisation("chq")]
        public ActionResult voucher()
        {
            return View(new VModel());
        }
        [DashboardAutherisation("chq")]
        public ActionResult cheque_vouchno()
        {
            return View(new VModel());
        }
        //[HttpPost]
        //public ActionResult chqdawn_post(Dashboard_New.Models.custom.chqdrawn v8, string grid)
        //{
        //    if (string.IsNullOrEmpty(v8.from_dt))
        //    {
        //        ModelState.AddModelError("from_dt", "Date Required");
        //    }
        //    if (string.IsNullOrEmpty(v8.to_dt))
        //    {
        //        ModelState.AddModelError("to_dt", "Date Required");
        //    }
        //    if (ModelState.IsValid)
        //    {                

        //        if (string.Equals("Submit", grid))
        //        {
        //            string from_dt = v8.from_dt;
        //            string to_dt = v8.to_dt;
        //            DateTime fromdt = DateTime.ParseExact(v8.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
        //            DateTime todt = DateTime.ParseExact(v8.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
        //            List<chqdrawn> records = new List<chqdrawn>();
        //            try
        //            {
        //                vcEntities vcobj = new vcEntities();
        //                records = vcobj.Database.SqlQuery<chqdrawn>(string.Format("select PARTY,BRNO,NPRNO,CQDATE,CHEQ_NO,RSAMT,VOUCHNO,CHEK  from CHECKDRAWN WHERE CQDATE >= '{0}' AND CQDATE<= '{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
        //                Dashboard_New.Models.VModel hg = new VModel();
        //                hg.chq = records;
        //                return View("cheque_vouchno", hg);
        //            }
        //            catch (Exception e)
        //            { Console.WriteLine("Erorrr : " + e); }
        //            VModel vd = new VModel();
        //            vd.chq = records;
        //            return View("cheque_vouchno", vd);
        //        }
        //    }
        //    else
        //    {
        //        return View("cheque_vouchno", v8);
        //    }
        //    return null;

        //}
        //[DashboardAutherisation("chqmdy")]
        //public ActionResult modify_cheques()
        //{
        //    return View(new VModel());
        //}
        public ActionResult Contact()
        { return View(); }

        [HttpPost]
        public FileResult Export(Dashboard_New.Models.VModel mod)
        {
            return null;
        }
        public ActionResult login()
        {
            return View();
        }
        [HttpPost]
        public ActionResult login(login l)
        {
            login lg = new login();
            if (ModelState.IsValid)
            {
                using (FACCTEntities lge = new FACCTEntities())
                {
                    var v = lge.useraccounts.Where(a => a.Name.Equals(l.UserName) && a.Password.Equals(l.Password)).FirstOrDefault();
                    if (v != null)
                    {
                        Session["LogedUserID"] = v.UserName.ToString();
                        Session["LogedUserFullname"] = v.Password.ToString();
                        return RedirectToAction("Index");
                    }
                }
            }
            return View("login");
        }
    }
}