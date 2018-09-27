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
    [DashboardAutherisation]
    public class HomeController : Controller
    {
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
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold=true;                
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
                        records = vcobj.Database.SqlQuery<vcvendor>(string.Format("select *, (select top 1 VName from vendormaster where VPartyCode = vvv.VPartyCode) VName from vcvouch vvv where DINP>= '{0}' AND DINP<= '{1}'", fromdt.ToString("yyyy-MM-dd",CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd",CultureInfo.InvariantCulture))).ToList();
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
                    var disp= entities.Database.SqlQuery<Dashboard_New.Models.custom.dcvendor>(string.Format("select DINP,DCTRNO,dbo.getAmount(AMOUNT) AS AMOUNT,NPRNO,ICCNO,COMNO,VRNO,BRNO,DCID,ASSTCK,ACCTCK,ACCT1CK,SOCK,DRCK,CRDATE,DCTRBNO,LUSER, (select top 1 COORNAME from DCMLST where DCID = vvv.DCID) VName from DVOUCH vvv where DINP>= '{0}' AND DINP<= '{1}'", fromdt.ToString("yyyy-MM-dd",CultureInfo.InvariantCulture),todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
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
        public ActionResult nirfpost(Dashboard_New.Models.custom.nirf v3,string grid, string export,string id)
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
                        wb.Worksheet(1).Cell(1, 7).Value = "NIRF(Sponsored) : R" + v3.from_year + " - "  + v3.to_year;
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
                        wb.Worksheet(1).Cell(1, 7).Value = "NIRF(Consultancy): R" + v3.from_year +" - " + v3.to_year;
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
                    { Console.WriteLine("Erorrr : " + e); }
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
                    dt.Columns.AddRange(new DataColumn[8] { new DataColumn("INDENT NUMBER"), new DataColumn("INDENT DATE"), new DataColumn("NAME"), new DataColumn("CURRENCY"), new DataColumn("INDENT VALUE"),new DataColumn("DEPARTMENT"), new DataColumn("STATUS"), new DataColumn("REMARKS") });
                     DateTime fromdt = DateTime.ParseExact(v4.from_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v4.to_dt, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.ppo>(string.Format("select indent_no,indent_date,[name],currency,indent_value,dept,[status],remarks from indent_master where indent_date>= '{0}' AND indent_date<= '{1}'", fromdt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), todt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))).ToList();
                    dt.Columns[4].DataType = typeof(Int32);
                    foreach (var x in disp)
                    {                        
                        dt.Rows.Add(x.indent_no, x.indent_date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),x.name,x.currency,x.indent_value, x.dept,x.status,x.remarks);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("Pending_Purchase_Order");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "Pending Indent : " + v4.from_dt +" To " + v4.to_dt;
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
                    string strddval = Request.Form["menu1"].ToString();
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


        public ActionResult pfms_vc()
        { return View(new VModel()); }

        [HttpPost]

        public ActionResult pfms_vc_post(Dashboard_New.Models.custom.pfms_vc v5, string grid, string export)
        {
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
                if (string.Equals("Export To Excel", export))
                {
                   
                    vcEntities entities = new vcEntities();
                    DataTable dt = new DataTable("Grid");
                    string from_dt = v5.from_year.Substring(2);
                    string to_dt = v5.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v5.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v5.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "VOU" + from_dt + to_dt;
                    dt.Columns.AddRange(new DataColumn[49] {new DataColumn("YEAR"), new DataColumn("DATE"), new DataColumn("AMOUNT"), new DataColumn("VRNO"), new DataColumn("NPRNO"), new DataColumn("PART"), new DataColumn("HEAD"), new DataColumn("DISC"), new DataColumn("DIS"), new DataColumn("ICCNO"), new DataColumn("PONO"), new DataColumn("COMNO"), new DataColumn("CQNO"), new DataColumn("BRNO"), new DataColumn("NATURE"), new DataColumn("CHECK"), new DataColumn("REGNO"),new DataColumn("LEDDIS"), new DataColumn("ECODE"), new DataColumn("VCTRNO"),new DataColumn("VPartyCode"), new DataColumn("ASSTCK"),new DataColumn("ACC1TCK"), new DataColumn("ACCTCK"), new DataColumn("SOCK"),new DataColumn("DRCK"), new DataColumn("CRDATE"), new DataColumn("CDSTATUS"), new DataColumn("TRANSFERED"), new DataColumn("EMAILID"), new DataColumn("VCTRBNO"), new DataColumn("LUSER"), new DataColumn("VName"), new DataColumn("VAddress"), new DataColumn("VPinCode"), new DataColumn("VMobile"), new DataColumn("VPhoneNumber"), new DataColumn("VEmailId"), new DataColumn("VPanNo"),new DataColumn("VTinNo"), new DataColumn("VserTaxRegNo"), new DataColumn("VAcctNameInBank"), new DataColumn("VBankName"),new DataColumn("VBranchName"), new DataColumn("VIFSCCode"), new DataColumn("VBankAcctNo"),new DataColumn("VBankMICRCode"), new DataColumn("VbankPhoneNumber"), new DataColumn("VBankEmailID") });
                    var disp = entities.Database.SqlQuery<Dashboard_New.Models.custom.pfms_vc>(string.Format("select v.DATE,v.AMOUNT,v.VRNO,v.NPRNO,v.PART,v.HEAD,v.DISC,v.DIS , V.ICCNO, V.PONO, V.COMNO, V.CQNO, V.BRNO, v.NATURE, v.[CHECK], v.REGNO,v.LEDDIS, v.ECODE, c.VCTRNO, c.VPartyCode, c.ASSTCK, c.ACCT1CK, c.ACCTCK,c.SOCK, c.DRCK, c.CRDATE, c.CDSTATUS, c.TRANSFERED, c.EMAILID, c.VCTRBNO, C.LUSER, M.VName,m.VAddress, m.VPinCode, m.VMobile, m.VPhoneNumber, m.VEmailId,m.VPanNo, m.VTinNo, m.VSerTaxRegNo, m.VAcctNameInBank, m.VBankName, m.VBranchName, m.VIFSCCode,M.VBankAcctNo, M.VBankName, M.VBankMICRCode, M.VBankPhoneNumber, M.VBankEmailID  from " + tabname + "  V,VENDORDRAWN C,VendorMaster M WHERE V.VRNO=C.VRNO AND C.VPartyCode=M.VPartyCode and v.nprno in (select nprno from mstlst where ACCOUNTTYPE='pfms')", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();

                    foreach (var x in disp)
                    {
                        dt.Rows.Add(x.DATE, x.AMOUNT, x.VRNO, x.NPRNO, x.PART,x.HEAD, x.DISC, x.DIS , x.ICCNO, x.PONO, x.COMNO, x.CQNO, x.BRNO, x.NATURE, x.CHECK, x.REGNO,x.LEDDIS, x.ECODE, x.VCTRNO, x.VPartyCode, x.ASSTCK, x.ACC1TCK, x.ACCTCK,x.SOCK, x.DRCK, x.CRDATE, x.CDSTATUS, x.TRANSFERED,x.EMAILID, x.VCTRBNO, x.LUSER, x.VName,x.VAddress, x.VPinCode, x.VMobile, x.VPhoneNumber, x.VEmailId,x.VPanNo, x.VTinNo, x.VserTaxRegNo, x.VAcctNameInBank, x.VBankName, x.VBranchName, x.VIFSCCode,x.VBankAcctNo, x.VBankMICRCode, x.VbankPhoneNumber, x.VBankEmailID);
                    }
                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        wb.Worksheets.Add("PFMS -VC");
                        wb.Worksheet(1).Cell(3, 1).InsertTable(dt);
                        wb.Worksheet(1).Cell(1, 7).Value = "PFMS (VC) : VOU" + v5.from_year + " - " + v5.to_year;
                        wb.Worksheet(1).Cell(1, 7).Style.Font.Bold = true;
                        var wbs = wb.Worksheets.FirstOrDefault();
                        wbs.Tables.FirstOrDefault().ShowAutoFilter = false;
                        wb.Properties.Title = "PFMS VC REPORT";
                        using (MemoryStream stream = new MemoryStream())
                        {
                            wb.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PFMS-VC.xlsx");
                        }
                    }
                }              
                if (string.Equals("Submit", grid))
                {
                    string from_dt = v5.from_year.Substring(2);
                    string to_dt = v5.to_year.Substring(2);
                    DateTime fromdt = DateTime.ParseExact(v5.from_year, "yyyy", CultureInfo.InvariantCulture);
                    DateTime todt = DateTime.ParseExact(v5.to_year, "yyyy", CultureInfo.InvariantCulture);
                    string tabname = "VOU" + from_dt + to_dt;
                    List<pfms_vc> records = new List<pfms_vc>();
                    try
                    {
                        vcEntities vcobj = new vcEntities();
                        records = vcobj.Database.SqlQuery<pfms_vc>(string.Format("select v.DATE,v.AMOUNT,v.VRNO,v.NPRNO,v.PART,v.HEAD,v.DISC,v.DIS , V.ICCNO, V.PONO, V.COMNO, V.CQNO, V.BRNO, v.NATURE, v.[CHECK], v.REGNO,v.LEDDIS, v.ECODE, c.VCTRNO, c.VPartyCode, c.ASSTCK, c.ACCT1CK, c.ACCTCK,c.SOCK, c.DRCK, c.CRDATE, c.CDSTATUS, c.TRANSFERED, c.EMAILID, c.VCTRBNO, C.LUSER, M.VName,m.VAddress, m.VPinCode, m.VMobile, m.VPhoneNumber, m.VEmailId,m.VPanNo, m.VTinNo, m.VSerTaxRegNo, m.VAcctNameInBank, m.VBankName, m.VBranchName, m.VIFSCCode,M.VBankAcctNo, M.VBankName, M.VBankMICRCode, M.VBankPhoneNumber, M.VBankEmailID  from "+ tabname + "  V,VENDORDRAWN C,VendorMaster M WHERE V.VRNO=C.VRNO AND C.VPartyCode=M.VPartyCode and v.nprno in (select nprno from mstlst where ACCOUNTTYPE='pfms')", fromdt.ToString("yyyy", CultureInfo.InvariantCulture), todt.ToString("yyyy", CultureInfo.InvariantCulture))).ToList();
                        Dashboard_New.Models.VModel dv = new VModel();
                        dv.from_year = v5.from_year;
                        dv.to_year = v5.to_year;
                        dv.pfm_vclist = records;
                        return View("pfms_vc", dv);
                    }
                    catch (Exception e)
                    { Console.WriteLine("Error : " + e); }
                    VModel vd = new VModel();
                    vd.pfm_vclist = records;
                    vd.from_year = v5.from_year;
                    vd.to_year = v5.to_year;
                    return View("pfms_vc", vd);
                }
               
            }
            else
            {
                return View("pfms_vc", v5);
            }
            return null;
        }
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