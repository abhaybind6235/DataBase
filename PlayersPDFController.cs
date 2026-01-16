using GameEvent.Models;
using GameEvent.DataBase;
using Document = iTextSharp.text.Document;
using Font = iTextSharp.text.Font;
using Image = iTextSharp.text.Image;
using Rectangle = iTextSharp.text.Rectangle;
using System.Web.Mvc;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Data;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using iTextSharp.text.pdf.draw;
using Aspose.Pdf.Operators;
using Org.BouncyCastle.Crypto.IO;
using Org.BouncyCastle.Utilities;
using System.Linq;
using Aspose.Pdf.Text;

namespace GameEvent.Controllers
{
    public class PlayersPDFController : Controller
    {
        DataLayer DbLayer = new DataLayer();
        dbHelper dbHelper = new dbHelper();
        CommonMethod Common = new CommonMethod();

        #region AttendanceSheet
        public ActionResult PlayerDetailsPDF(int Evt_Id, int OrgId)
        {
            Players model = new Players();
            model.EventId = Evt_Id;
            model.OrgId = OrgId;

            model.dt = DbLayer.DisplayLists(43, OrgId, Evt_Id);
            model.dt1 = DbLayer.DisplayLists(5, OrgId, Evt_Id);
            model.dt5 = DbLayer.DisplayLists(57, OrgId, Evt_Id);
            model.dt2 = DbLayer.DisplayLists(24, Evt_Id);
            return GeneratePDF(model);
        }


        private ActionResult GeneratePDF(Players model)
        {
            GENERATE_DOWNLOAD_History model1 = new GENERATE_DOWNLOAD_History();

            model1.Action = 3;
            model1.Event_Id = model.EventId;
            model1.Org_Id = model.OrgId;
            model1.Doc_Status = "GeneratePlayer";
            model1.CreatedBy = int.Parse(Session["UserId"].ToString());
            model.dt4 = DbLayer.Attendance_Gen_History(model1);
            if (model.dt4.Rows.Count == 0)
            {
                using (MemoryStream memStream = new MemoryStream())
                {
                    Document doc = new Document(PageSize.A4);
                    try
                    {
                        string footer = model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString();
                        PdfWriter writer = PdfWriter.GetInstance(doc, memStream);
                        writer.PageEvent = new PDFFooter(footer);
                        doc.Open();

                        AddTitle(doc, model.dt2);
                        AddPlayerTable(doc, model);
                        doc.NewPage();
                        AddOfficialTable(doc, model);
                        AddAffiliatedUnitsAgreement(doc);

                        doc.Close();

                        byte[] bytes = memStream.ToArray();
                        string fileName = DateTime.Now.Ticks + "_PlayerDetailsPDF.pdf";
                        string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), fileName);
                        System.IO.File.WriteAllBytes(fullPath, bytes);


                        model1.Action = 1;
                        model1.Doc_Type = "PlayerDetailsPDF";
                        model1.Doc_Status = "GeneratePlayer";
                        model1.DocumentName = fileName;
                        model.dt4 = DbLayer.Attendance_Gen_History(model1);


                        return File(bytes, "application/pdf", fileName);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                        return new HttpStatusCodeResult(500, "Internal Server Error");
                    }
                }
            }
            else
            {
                model1.Action = 1;
                model1.Doc_Type = "PlayerDetailsPDF";
                model1.Doc_Status = "DownloadPlayer";
                model1.DocumentName = model.dt4.Rows[0][1].ToString();
                model.dt4 = DbLayer.Attendance_Gen_History(model1);
                string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), model1.DocumentName);
                return File(fullPath, "application/pdf", model1.DocumentName);
            }
        }
        private void AddTitle(Document doc, DataTable dt)
        {
            Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 20, new BaseColor(163, 67, 123));
            Paragraph title = new Paragraph("VOLLEYBALL FEDERATION OF INDIA", titleFont)
            {
                Alignment = Element.ALIGN_CENTER
            };
            doc.Add(title);
            doc.Add(new Paragraph("\n"));

            PdfPTable headerTable = new PdfPTable(3) { WidthPercentage = 100 };
            headerTable.SetWidths(new float[] { 30f, 40f, 30f });

            headerTable.AddCell(CreateCell(dt.Rows[0]["GameName"].ToString() + " " + dt.Rows[0]["AgeGroupName"].ToString() + " " + dt.Rows[0]["Gender"].ToString(), Element.ALIGN_LEFT));
            headerTable.AddCell(CreateCell(dt.Rows[0]["Calendar"].ToString() + "\n" + dt.Rows[0]["Session"].ToString(), Element.ALIGN_CENTER, true));
            headerTable.AddCell(CreateCell("State: " + Session["OrganizatinShortName"].ToString(), Element.ALIGN_RIGHT));

            doc.Add(headerTable);

            PdfPTable eventTable = new PdfPTable(1) { WidthPercentage = 100 };


            StringBuilder details = new StringBuilder();
            details.AppendLine(dt.Rows[0]["StartDate"].ToString() + " To " + dt.Rows[0]["EndDate"].ToString() + "\n");
            details.AppendLine("Organized By:" + dt.Rows[0]["OrganizatinName"].ToString() + "\n");
            details.AppendLine("Under the aegis of Volleyball Federation of India");


            //string details = "11.11.2024 To 12.11.2024, Nadiad\nOrganized By: Sports Authority of Gujarat, Gandhinagar\nUnder the aegis of School Games Federation of India";
            eventTable.AddCell(CreateCell(details.ToString(), Element.ALIGN_CENTER));
            doc.Add(eventTable);

            PdfPTable headerTable2 = new PdfPTable(2) { WidthPercentage = 100 };
            headerTable2.SetWidths(new float[] { 75f, 25f });
            headerTable2.AddCell(CreateCell(dt.Rows[0]["Calendar"].ToString() + " " + dt.Rows[0]["Session"].ToString(), Element.ALIGN_RIGHT, true));
            headerTable2.AddCell(CreateCell("Date: " + DateTime.Now.ToString("dd/MMM/yyyy"), Element.ALIGN_RIGHT));
            doc.Add(headerTable2);
        }
        private void AddPlayerTable(Document doc, Players model)
        {
            PdfPTable table = new PdfPTable(9) { WidthPercentage = 100 };

            PdfPCell fullWidthCell = new PdfPCell(new Phrase(model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString(), FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.WHITE)))
            {
                Colspan = 9,
                HorizontalAlignment = Element.ALIGN_CENTER,
                Border = Rectangle.BOX,
                BackgroundColor = new BaseColor(0, 102, 204),
                PaddingBottom = 8
            };
            table.AddCell(fullWidthCell);
            table.SetWidths(new float[] { 5f, 15f, 15f, 15f, 10f, 10f, 15f, 10f, 10f });

            string[] headers = { "SN", "Reg. No", "Name", "Father's Name", "DOB", "Cls", "Category", "School Name", "Photo" };
            foreach (string header in headers)
                table.AddCell(CreateHeaderCell(header));

            int i = 1;
            Certificate Certi = new Certificate();
            Certi.Action = 1;
            Certi.CertificateTypeId = 1;
            Certi.CertificateIssuedBy = Convert.ToInt32(Session["UserId"]);
            foreach (DataRow player in model.dt.Rows)
            {
                Certi.CertificatePlayerId = player["PlayerId"].ToString();
                Certi.CertificateTeamRankId = model.dt5.Rows.Count > 0 ? Convert.ToInt32(model.dt5.Rows[0]["TeamRank_Id"]) : 0;
                var data = DbLayer.CertificateManage(Certi);
                table.AddCell(CreateBodyCell((i++).ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["SGFIRegNo"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Name"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["FatherName"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(Convert.ToDateTime(player["DateOfBirth"]).ToString("dd/MM/yyyy"), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["PlayerClass"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["GameName"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["SchoolName"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(GetPlayerImageCell(player["PlayerPhotograph"].ToString()));

            }
            doc.Add(table);
            doc.Add(new Paragraph("\n"));
        }

        private void AddOfficialTable(Document doc, Players model)
        {
            PdfPTable table = new PdfPTable(7) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 5f, 20f, 20f, 15f, 15f, 15f, 10f });
            PdfPCell fullWidthCell = new PdfPCell(new Phrase("COACH / MANAGER / HOD (" + model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString() + ")", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, BaseColor.WHITE)))
            {
                Colspan = 7,
                HorizontalAlignment = Element.ALIGN_CENTER,
                Border = Rectangle.BOX,
                BackgroundColor = new BaseColor(0, 102, 204),
                PaddingBottom = 8
            };
            table.AddCell(fullWidthCell);

            string[] headers = { "SN", "Name", "Father's Name", "Designation", "DOB", "Mobile No", "Photo" };
            foreach (string header in headers)
                table.AddCell(CreateHeaderCell(header));

            int i = 1;
            foreach (DataRow player in model.dt1.Rows)
            {
                table.AddCell(CreateBodyCell((i++).ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Name"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["FatherName"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Type"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(Convert.ToDateTime(player["DateOfBirth"]).ToString("dd/MM/yyyy"), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Phone"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(GetPlayerImageCell(player["Photograph1"].ToString()));
            }
            doc.Add(table);
        }
        private void AddAffiliatedUnitsAgreement(Document doc)
        {
            Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.BLACK);
            Font bodyFont = FontFactory.GetFont(FontFactory.HELVETICA, 10, BaseColor.BLACK);
            Font textFont = FontFactory.GetFont(FontFactory.HELVETICA, 8, BaseColor.BLACK);

            Paragraph title = new Paragraph("Affiliated Unit's Agreement", titleFont)
            {
                Alignment = Element.ALIGN_CENTER
            };
            doc.Add(title);
            doc.Add(new Paragraph("\n"));

            StringBuilder agreementText = new StringBuilder();
            agreementText.AppendLine(" I hereby agree to the following terms & conditions governing the ontine official entry from process of participation in 68th National School Games (2024-25) auspices in S.G.F.I. \n");
            agreementText.AppendLine("1. I have gone through and understood the contents of information brochure and eligibility criteria prescribed there in  I shall abide by rules & regulation of online official entry form process for participation in 68th National School Games (2024-25) as specified in the information brochure, rules & regulation of organization of National School Games .");
            agreementText.AppendLine("\n l am aware of fact that in any unavoidable circumstances SGFI shall have the power to prepone, postpone or cancel the said National School Games and SGFl is not liable for any financial loss occured in this situation.  ");

            agreementText.AppendLine("\n 2. I certify that the eligibility of players who's official entry form is being filled has been fulfilled according to the rules of championship and also certify that these players are students of class 12th or below class but not below 6th Class. ");
            agreementText.AppendLine("\n3. I declare that each one of our team's players are born on or after date 01-01-2008 and hence they are eligible to participate in their respective age group. I declare that in this calendar year 2024-25. The player mentioned above in the official entry form shall participate in discipline as mentioned above only through the specified age group in the official entry form and shall not participate through any other age group in same discipline. ");
            agreementText.AppendLine("\n 4. I know that during verification of documents at the time of reporting, if any discrepancy is detected in original document including name, father's name, date of birth, class, admission number, school name, eligibility and gender, then my unit's team / subjected players participation will be liable to be cancelled.  ");

            agreementText.AppendLine("\n 5. I declare that I am aware that in the process of filling online entry form, after me completing the entry & once | click on confirm & print button, then I cannot make any correction at my end. This would be considered as final entry and accordingly S.G.F.I. will make the participation of such players and on the base of this the ldentity Card, Participation Certificate, Merit Certificate will be issued by S.G.F.I.   ");
            agreementText.AppendLine("\n 6. I declare that the information provided by me is genuíne & authentic.");
            agreementText.AppendLine("\n7. I declare that for the participation in National School Championship in which our team is participating we have fulfilled all the required eligibility criteria for age & class in which the candidate is a regular student in recognized schools in our unit administration.   ");
            agreementText.AppendLine("\n8. I declare that I will not disclose or share the password & event code provided by SGFI for process of filling online official entry form of above said championship with anybody. I understand that I am solely responsible for safeguarding password & event code and S.G.F.I. is not responsible for any misuse of my password & event code.");
            agreementText.AppendLine("\n9.I declare that, I shall be responsible for the safety & comfort of players of my team during their travel from their home to the venue of tournament and back to their home. The travel ticket expenses and the expenses during travel will be borne by our unit. The food expense shall be completely borne by our unit during the tournament including travel period. Our unit shall bear the expenses of kit, dress, etc provided to the players, Our unit shall provide medical facilities to our team's players. Our unit will also provide medical & accidental insurance to our team's each and every player & member. ");
            agreementText.AppendLine("\n10. I declare that l am very well aware of the fact that if any plaver is absent in tournament after filling the online official entry form then it has to be intimated to organizer in writing at control room one day before the tournament. If I fail to do so then I shall be liable for legal action under fraud case.  ");
            agreementText.AppendLine("\nl am aware that I have to take the signature of all the players mentioned above in the attendance sheet provided Dy SGFI office before participating in the said national school tournament and further the attendance sheet shall also De counter signed by our unit,s coach & chief-de-mission and submit the same in the control room situated at the venue of tournament one day before the tournament. ");
            agreementText.AppendLine("\n11. I hereby declare that I shall submit the following documents along with  the printout of the official entry form(Triplicate) which will be obtained after clicking on the confirm & print button. The print out of official entry form shall be signed by the competent authority, coach & manager and with three copies of the same along with the following documents shall be submitted by chief-de-mission in the control room situated at venue of tournament mandatorily before one day prior to National School Tournament: ");
            agreementText.AppendLine("\n12. Online Official Entries for said tournament will be closed 5 days before tournament dates. The Final Print Out before 5 days will be treated as final official entry form. ");
            agreementText.AppendLine("\n13. Rs. 200 will be charged for issue of every duplicate Participation/ Merit certificate of any National School Games auspices in SGFI. Payment should be made in favour of SCHOOL GAMES FEDERATION OF INDIA, payable by concerning STATE/UT/UNIT.\n");

            doc.Add(new Paragraph(agreementText.ToString(), textFont));


            Paragraph paragraph1 = new Paragraph();
            paragraph1.Add(new Chunk("1. Covering letter with authority letter (which is obtained after clicking confirm & print button along with official entry form): ",
                                    FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));
            paragraph1.Add(new Chunk("Signed by the competent authority whose specimen signature was sent to S.G.F.I. office by the competent authority at the time of annual recognition.",
                                    FontFactory.GetFont(FontFactory.HELVETICA, 8)));
            paragraph1.Add(new Chunk(" - (One)\n", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));

            paragraph1.Add(new Chunk("2. Eligibility certificate: ", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));
            paragraph1.Add(new Chunk("Separate eligibility certificates of each & every player of the team issued by the school in which the player is studying regularly and should have the signature of Principle / Head Master of school and further countersigned by the competent authority of the unit/Authorized officer.",
                                    FontFactory.GetFont(FontFactory.HELVETICA, 8)));
            paragraph1.Add(new Chunk(" - (Triplicate) \n", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));

            paragraph1.Add(new Chunk("3. Birth certificate:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));
            paragraph1.Add(new Chunk("Separate photo-copy of Birth certificates of each & every player of the team attested by Gazetted officer. Only the birth certificate issued by the Statistic Department of state/UT Govt. /Central Govt. or Municipal Corporation shall be acceptable",
                                    FontFactory.GetFont(FontFactory.HELVETICA, 8)));
            paragraph1.Add(new Chunk(" - (Triplicate) \n", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));

            paragraph1.Add(new Chunk("4. Previous year final exam mark sheet:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));
            paragraph1.Add(new Chunk("Separate photo-copy of Mark sheet of each & every player of the team, attested by the Gazetted officer.",
                                    FontFactory.GetFont(FontFactory.HELVETICA, 8)));
            paragraph1.Add(new Chunk(" - (Triplicate) \n", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));


            paragraph1.Add(new Chunk("5. Registration + certificate ID card fees:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));
            paragraph1.Add(new Chunk("@ Rs. 300/- per player.\n", FontFactory.GetFont(FontFactory.HELVETICA, 8)));

            paragraph1.Add(new Chunk(" 6. Copy of AADHAAR Card  \n", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));
            paragraph1.Add(new Chunk("l am aware that the participation of our above mentioned team will be confirmed only when the above mentioned documents dully signed and  registration fees is submitted at the control room situated at the venue of the tournament one day before the National School Tournament and  subject to validation of submitted documents.  \n \n", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));
            paragraph1.Add(new Chunk("I hereby declare that I agree to all the terms & conditions of above said agreement from point number 1 to 13. . \n", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8)));

            doc.Add(paragraph1);
            PdfPTable sign = new PdfPTable(1) { WidthPercentage = 100 };
            sign.SetWidths(new float[] { 100f });
            sign.AddCell(CreateCell("Signature of Competent Authority of State/UT/Unit.\n\n", Element.ALIGN_CENTER, fontsize: 10));
            doc.Add(sign);


            PdfPTable headerTable = new PdfPTable(3) { WidthPercentage = 100 };
            headerTable.SetWidths(new float[] { 33f, 33f, 34f });

            headerTable.AddCell(CreateCell("Signature of Coach:................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Signature of Manager:..............", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("(With official Seal)\n", Element.ALIGN_LEFT, fontsize: 10));

            headerTable.AddCell(CreateCell("Name:...................................\n", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Name:...................................\n", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Name:...................................\n", Element.ALIGN_LEFT, fontsize: 10));

            headerTable.AddCell(CreateCell("Original Post:...........................\n", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Original Post:...........................\n", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Original Post:...........................\n", Element.ALIGN_LEFT, fontsize: 10));

            headerTable.AddCell(CreateCell("Dept Address:........................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Dept Address:........................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Dept Address:........................", Element.ALIGN_LEFT, fontsize: 10));

            headerTable.AddCell(CreateCell("Mobile No:.............................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Mobile No:.............................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Mobile No:.............................", Element.ALIGN_LEFT, fontsize: 10));

            headerTable.AddCell(CreateCell("Date & Place:....................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("", Element.ALIGN_LEFT, fontsize: 10));
            doc.Add(headerTable);

            PdfPTable AUTHORITY = new PdfPTable(1) { WidthPercentage = 100 };
            AUTHORITY.SetWidths(new float[] { 100f });
            Phrase phrase = new Phrase("\n\n\n\nAUTHORITY\n", FontFactory.GetFont(FontFactory.HELVETICA, 12, Font.UNDERLINE));
            PdfPCell cell = new PdfPCell(phrase)
            {
                Border = PdfPCell.NO_BORDER,
                HorizontalAlignment = Element.ALIGN_CENTER
            };
            AUTHORITY.AddCell(cell);
            doc.Add(AUTHORITY);


            Paragraph paragraph2 = new Paragraph();
            paragraph2.Add(new Chunk("It is certify that Sh________________________Post______________________office address______________________________mobile no.__________________ is being authorized from our unit for sign of documents related above mentioned tournament. The specimen signature of our above mentioned representative is certified, which is given below. Specimen signature of authorized officer ...............................................................kinidly allow our team to participate in the said tournaments.",
                                    FontFactory.GetFont(FontFactory.HELVETICA, 8)));
            paragraph2.Add(new Chunk("\nThanking you \n", FontFactory.GetFont(FontFactory.HELVETICA, 8)));

            doc.Add(paragraph2);


            PdfPTable AUTH = new PdfPTable(2) { WidthPercentage = 100 };
            AUTH.SetWidths(new float[] { 50f, 50f });

            AUTH.AddCell(CreateCell("Date.............", Element.ALIGN_CENTER, fontsize: 10));
            AUTH.AddCell(CreateCell("Seal and Signature of Competent Authority", Element.ALIGN_CENTER, fontsize: 10));

            AUTH.AddCell(CreateCell("Place..........", Element.ALIGN_CENTER, fontsize: 10));
            AUTH.AddCell(CreateCell("(Competent Authority of Affiliated Unit)", Element.ALIGN_CENTER, fontsize: 10));

            AUTH.AddCell(CreateCell("", Element.ALIGN_CENTER, fontsize: 10));
            AUTH.AddCell(CreateCell("Name", Element.ALIGN_CENTER, fontsize: 10));

            AUTH.AddCell(CreateCell("", Element.ALIGN_CENTER, fontsize: 10));
            AUTH.AddCell(CreateCell("Post", Element.ALIGN_CENTER, fontsize: 10));

            AUTH.AddCell(CreateCell("", Element.ALIGN_CENTER, fontsize: 10));
            AUTH.AddCell(CreateCell("Office address", Element.ALIGN_CENTER, fontsize: 10));

            AUTH.AddCell(CreateCell("", Element.ALIGN_CENTER, fontsize: 10));
            AUTH.AddCell(CreateCell("Mobile/ land line", Element.ALIGN_CENTER, fontsize: 10));

            doc.Add(AUTH);
        }


        private PdfPCell CreateCell(string text, int alignment, bool isBold = false, int fontsize = 8)
        {
            Font font = isBold ? FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10) : FontFactory.GetFont(FontFactory.HELVETICA, 10);
            return new PdfPCell(new Phrase(text, font)) { Border = Rectangle.NO_BORDER, HorizontalAlignment = alignment };
        }


        private PdfPCell CreateBodyCell(string text, int alignment)
        {
            Font font = FontFactory.GetFont(FontFactory.HELVETICA, 10);
            return new PdfPCell(new Phrase(text, font))
            {
                HorizontalAlignment = alignment,
                Border = Rectangle.BOX,
                Padding = 5
            };
        }
        private PdfPCell CreateHeaderCell1(string text)
        {
            Font font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK);
            return new PdfPCell(new Phrase(text, font))
            {
                BackgroundColor = BaseColor.WHITE,
                HorizontalAlignment = Element.ALIGN_CENTER,
                Padding = 5
            };
        }

        private PdfPCell CreateHeaderCell(string text)
        {
            Font font = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK);
            return new PdfPCell(new Phrase(text, font))
            {
                BackgroundColor = new BaseColor(200, 200, 200),
                HorizontalAlignment = Element.ALIGN_CENTER,
                Padding = 5
            };
        }

        private PdfPCell GetPlayerImageCell(string imagePath)
        {
            try
            {
                string fullPath = HttpContext.Server.MapPath(".." + imagePath);
                if (System.IO.File.Exists(fullPath))
                {
                    Image img = Image.GetInstance(fullPath);
                    img.ScaleAbsolute(40, 40);
                    return new PdfPCell(img) { HorizontalAlignment = Element.ALIGN_CENTER, Border = Rectangle.BOX, Padding = 5 };
                }
            }
            catch
            {

            }
            PdfPCell noPhotoCell = CreateBodyCell("No Photo", Element.ALIGN_CENTER);
            noPhotoCell.Border = Rectangle.BOX;
            return noPhotoCell;
        }

        static Paragraph MakeBoldText(string boldText)
        {
            Font boldFont1 = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
            Paragraph paragraph = new Paragraph(new Chunk(boldText, boldFont1));
            return paragraph;
        }

        public class PdfWatermarkHelper : PdfPageEventHelper
        {
            public override void OnEndPage(PdfWriter writer, Document document)
            {
                PdfContentByte canvas = writer.DirectContentUnder;
                BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED);
                canvas.BeginText();
                canvas.SetFontAndSize(baseFont, 10);
                canvas.SetColorFill(new GrayColor(0.95f));

                // Watermark multiple times on each page
                for (float x = 10; x < document.PageSize.Width; x += 155)
                {
                    for (float y = -50; y < document.PageSize.Height + 100; y += 10)
                    {
                        canvas.ShowTextAligned(Element.ALIGN_CENTER, "VOLLEYBALL FEDERATION OF INDIA", x, y, 30);
                    }
                }

                canvas.EndText();
            }
        }

        public class PDFFooter : PdfPageEventHelper
        {
            private string footerText;

            public PDFFooter(string footerText)
            {
                this.footerText = footerText;
            }

            public override void OnEndPage(PdfWriter writer, Document document)
            {
                PdfPTable footer = new PdfPTable(2);
                footer.TotalWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                footer.SetWidths(new float[] { 85f, 15f });

                // Horizontal Line
                PdfContentByte cb = writer.DirectContent;
                cb.MoveTo(document.LeftMargin, document.BottomMargin);
                cb.LineTo(document.PageSize.Width - document.RightMargin, document.BottomMargin);
                cb.Stroke();

                // Centered Content
                PdfPCell leftCell = new PdfPCell(new Phrase(footerText, FontFactory.GetFont(FontFactory.HELVETICA, 8)));
                leftCell.HorizontalAlignment = Element.ALIGN_CENTER;
                leftCell.Border = Rectangle.NO_BORDER;
                leftCell.PaddingTop = 6f;

                // Page Number on Right
                PdfPCell rightCell = new PdfPCell(new Phrase("Page " + writer.PageNumber, FontFactory.GetFont(FontFactory.HELVETICA, 8)));
                rightCell.HorizontalAlignment = Element.ALIGN_RIGHT;
                rightCell.Border = Rectangle.NO_BORDER;
                rightCell.PaddingTop = 6f;

                footer.AddCell(leftCell);
                footer.AddCell(rightCell);

                footer.WriteSelectedRows(0, -1, document.LeftMargin, document.BottomMargin, writer.DirectContent);
            }
        }
        #endregion


        #region Admin AttendanceSheet
        public ActionResult AllAttendanceSheetPDF(int Evt_Id)
        {
            try
            {
                Players model = new Players();
                GENERATE_DOWNLOAD_History model1 = new GENERATE_DOWNLOAD_History();
                model.dt2 = DbLayer.DisplayLists(24, Evt_Id);
                model1.Action = 2;
                model1.Event_Id = Evt_Id;
                model1.Doc_Type = "AllAttendanceSheetPDF";
                model1.Doc_Status = "AllGenerate";
                model1.CreatedBy = int.Parse(Session["UserId"].ToString());
                model.dt4 = DbLayer.Attendance_Gen_History(model1);
                if (model.dt4.Rows.Count == 0)
                {
                    using (MemoryStream memStream = new MemoryStream())
                    {
                        Document doc = new Document(PageSize.A4.Rotate());

                        model.dt3 = DbLayer.DisplayLists(44, EventId: Evt_Id);
                        if (model.dt3.Rows.Count > 0)
                        {
                            string footer = model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString();
                            PdfWriter writer = PdfWriter.GetInstance(doc, memStream);
                            writer.PageEvent = new PDFFooter(footer);
                            doc.Open();
                            foreach (DataRow item in model.dt3.Rows)
                            {
                                var data11 = dbHelper.ExecAdaptorDataTable("select OrganizatinName,OrganizatinShortName from M_Organizatins where IsActive=1 and OrganizatinId=" + Convert.ToInt32(item["Org_Id"]));
                                if (data11.Rows.Count > 0)
                                {
                                    model.OrganizatinShortName = data11.Rows[0]["OrganizatinShortName"].ToString();
                                    model.OrganizatinName = data11.Rows[0]["OrganizatinName"].ToString();
                                }
                                model.dt = DbLayer.DisplayLists(43, Convert.ToInt32(item["Org_Id"]), Evt_Id);
                                model.dt1 = DbLayer.DisplayLists(5, Convert.ToInt32(item["Org_Id"]), Evt_Id);
                                model.dt5 = DbLayer.DisplayLists(57, Convert.ToInt32(item["Org_Id"]), Evt_Id);

                                PlayersAttendanceTitle(doc, model);
                                AttendanceContent(doc, model);
                                doc.NewPage();
                                OfficialAttendanceTitle(doc, model);
                                AttendanceContent(doc, model);
                                doc.Add(new Paragraph("\n\n\n\n"));
                            }
                        }
                        doc.Close();

                        byte[] bytes = memStream.ToArray();
                        string fileName = DateTime.Now.Ticks + "_AttendanceSheet.pdf";
                        string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), fileName);
                        System.IO.File.WriteAllBytes(fullPath, bytes);


                        model1.Action = 1;
                        model1.Doc_Status = "AllGenerate";
                        model1.DocumentName = fileName;
                        model.dt4 = DbLayer.Attendance_Gen_History(model1);


                        return File(bytes, "application/pdf", fileName);
                    }
                }
                else
                {
                    model1.Action = 1;
                    model1.Doc_Status = "AllDownload";
                    model1.DocumentName = model.dt4.Rows[0][1].ToString();
                    model.dt4 = DbLayer.Attendance_Gen_History(model1);
                    string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), model1.DocumentName);
                    return File(fullPath, "application/pdf", model1.DocumentName);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                return new HttpStatusCodeResult(500, "Internal Server Error");
            }
        }
        public ActionResult AttendanceSheetPDF(int Evt_Id, int OrgId)
        {
            Players model = new Players();
            model.EventId = Evt_Id;
            model.OrgId = OrgId;
            var data11 = dbHelper.ExecAdaptorDataTable("select OrganizatinName,OrganizatinShortName from M_Organizatins where IsActive=1 and OrganizatinId=" + OrgId);
            if (data11.Rows.Count > 0)
            {
                model.OrganizatinShortName = data11.Rows[0]["OrganizatinShortName"].ToString();
                model.OrganizatinName = data11.Rows[0]["OrganizatinName"].ToString();
            }

            model.dt = DbLayer.DisplayLists(43, OrgId, Evt_Id);
            model.dt1 = DbLayer.DisplayLists(5, OrgId, Evt_Id);
            model.dt5 = DbLayer.DisplayLists(57, OrgId, Evt_Id);

            model.dt2 = DbLayer.DisplayLists(24, Evt_Id);
            return GenerateAttendanceSheetPDF(model);
        }
        private ActionResult GenerateAttendanceSheetPDF(Players model)
        {
            GENERATE_DOWNLOAD_History model1 = new GENERATE_DOWNLOAD_History();

            model1.Action = 3;
            model1.Event_Id = model.EventId;
            model1.Org_Id = model.OrgId;
            model1.Doc_Type = "AttendanceSheetPDF";
            model1.Doc_Status = "Generate";
            model1.CreatedBy = int.Parse(Session["UserId"].ToString());
            model.dt4 = DbLayer.Attendance_Gen_History(model1);
            if (model.dt4.Rows.Count == 0)
            {
                using (MemoryStream memStream = new MemoryStream())
                {
                    Document doc = new Document(PageSize.A4.Rotate());
                    try
                    {
                        string footer = model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString();
                        PdfWriter writer = PdfWriter.GetInstance(doc, memStream);
                        writer.PageEvent = new PDFFooter(footer);
                        doc.Open();

                        PlayersAttendanceTitle(doc, model);
                        AttendanceContent(doc, model);
                        doc.NewPage();
                        OfficialAttendanceTitle(doc, model);
                        AttendanceContent(doc, model);

                        doc.Close();

                        byte[] bytes = memStream.ToArray();
                        string fileName = DateTime.Now.Ticks + "_AttendanceSheet.pdf";
                        string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), fileName);
                        System.IO.File.WriteAllBytes(fullPath, bytes);


                        model1.Action = 1;
                        model1.Doc_Type = "AttendanceSheetPDF";
                        model1.Doc_Status = "Generate";
                        model1.DocumentName = fileName;
                        model.dt4 = DbLayer.Attendance_Gen_History(model1);


                        return File(bytes, "application/pdf", fileName);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                        return new HttpStatusCodeResult(500, "Internal Server Error");
                    }
                }
            }
            else
            {
                model1.Action = 1;
                model1.Doc_Type = "AttendanceSheetPDF";
                model1.Doc_Status = "Download";
                model1.DocumentName = model.dt4.Rows[0][1].ToString();
                model.dt4 = DbLayer.Attendance_Gen_History(model1);
                string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), model1.DocumentName);
                return File(fullPath, "application/pdf", model1.DocumentName);
            }
        }

        private void PlayersAttendanceTitle(Document doc, Players model)
        {
            PdfPTable titleTable = new PdfPTable(3) { WidthPercentage = 100 };
            titleTable.SetWidths(new float[] { 33f, 34f, 33f });
            PdfPCell leftCell = new PdfPCell(new Phrase("STATE/UNIT WISE:-" + model.dt2.Rows[0]["OrganizatinName"].ToString(), FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9)))
            {
                HorizontalAlignment = Element.ALIGN_LEFT,
                Border = Rectangle.TOP_BORDER | Rectangle.LEFT_BORDER,
                Padding = 8
            };
            titleTable.AddCell(leftCell);

            PdfPCell centerCell = new PdfPCell();
            centerCell.Border = Rectangle.TOP_BORDER;

            Paragraph centerParagraph = new Paragraph();
            centerParagraph.Add(new Chunk(model.dt2.Rows[0]["Calendar"].ToString() + "\n", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12)));
            centerParagraph.SpacingBefore = 10f;
            centerParagraph.SetLeading(15f, 0); // 50px space between lines
            centerParagraph.Add(new Chunk(model.dt2.Rows[0]["StartDate"].ToString() + " To " + model.dt2.Rows[0]["EndDate"].ToString() + "(" + model.dt2.Rows[0]["Venue"].ToString() + "- " + model.dt2.Rows[0]["OrganizatinShortName"].ToString() + ")", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)));
            centerCell.AddElement(centerParagraph);

            centerCell.HorizontalAlignment = Element.ALIGN_CENTER;
            centerCell.Padding = 8;
            titleTable.AddCell(centerCell);

            PdfPCell rightCell = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            {
                Border = Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER,
                Padding = 8
            };
            titleTable.AddCell(rightCell);

            doc.Add(titleTable);



            PdfPTable table = new PdfPTable(12) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 5f, 15f, 15f, 20f, 12f, 7f, 20f, 10f, 8f, 10f, 15f, 10f });

            // **Game/Unit Row**
            PdfPCell GameCell = new PdfPCell(new Phrase(model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " " + model.dt2.Rows[0]["OrganizatinShortName"].ToString(), FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, Font.ITALIC)))
            {
                Colspan = 12,
                HorizontalAlignment = Element.ALIGN_CENTER,
                Border = Rectangle.BOX,
                Padding = 5
            };
            table.AddCell(GameCell);


            string[] headers = { "Sr No.", "Reg No.", "NAME", "FATHER NAME", "DOB", "CLASS", "SCHOOL", "ADM. NO", "GAME CATEGORY", "PHOTO", "SIGNATURE", "CONTACT NO." };

            foreach (string header in headers)
                table.AddCell(CreateHeaderCell1(header));

            // **State/Unit Row**
            PdfPCell stateCell = new PdfPCell(new Phrase("State/Unit-" + model.OrganizatinShortName, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, Font.ITALIC)))
            {
                Colspan = 12,
                HorizontalAlignment = Element.ALIGN_LEFT,
                Border = Rectangle.BOX,
                Padding = 5
            };
            table.AddCell(stateCell);

            int i = 1;
            Certificate Certi = new Certificate();
            Certi.Action = 1;
            Certi.CertificateTypeId = 1;
            Certi.CertificateIssuedBy = Convert.ToInt32(Session["UserId"]);
            foreach (DataRow player in model.dt.Rows)
            {
                Certi.CertificatePlayerId = player["PlayerId"].ToString();
                Certi.CertificateTeamRankId = model.dt5.Rows.Count > 0 ? Convert.ToInt32(model.dt5.Rows[0]["TeamRank_Id"]) : 0;
                var data = DbLayer.CertificateManage(Certi);

                table.AddCell(CreateBodyCell((i++).ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["SGFIRegNo"]?.ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Name"]?.ToString(), Element.ALIGN_LEFT));
                table.AddCell(CreateBodyCell(player["FatherName"]?.ToString(), Element.ALIGN_LEFT));

                string dob = DateTime.TryParse(player["DateOfBirth"]?.ToString(), out DateTime dateOfBirth)
                             ? dateOfBirth.ToString("dd-MM-yyyy")
                             : "N/A";
                table.AddCell(CreateBodyCell(dob, Element.ALIGN_CENTER));

                table.AddCell(CreateBodyCell(player["PlayerClass"]?.ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["SchoolName"]?.ToString(), Element.ALIGN_LEFT));
                table.AddCell(CreateBodyCell(player["AdmissionNo"]?.ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["GameName"]?.ToString(), Element.ALIGN_CENTER));

                table.AddCell(GetPlayerImageCell(player["PlayerPhotograph"]?.ToString()));
                table.AddCell(CreateBodyCell("", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("", Element.ALIGN_CENTER));
            }

            doc.Add(table);
            doc.Add(new Paragraph("\n"));
        }
        private void OfficialAttendanceTitle(Document doc, Players model)
        {
            PdfPTable titleTable = new PdfPTable(3) { WidthPercentage = 100 };
            titleTable.SetWidths(new float[] { 33f, 34f, 33f });
            PdfPCell leftCell = new PdfPCell(new Phrase("STATE/UNIT WISE:-" + model.dt2.Rows[0]["OrganizatinName"].ToString(), FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9)))
            {
                HorizontalAlignment = Element.ALIGN_LEFT,
                Border = Rectangle.TOP_BORDER | Rectangle.LEFT_BORDER,
                Padding = 8
            };
            titleTable.AddCell(leftCell);

            PdfPCell centerCell = new PdfPCell();
            centerCell.Border = Rectangle.TOP_BORDER;

            Paragraph centerParagraph = new Paragraph();
            centerParagraph.Add(new Chunk(model.dt2.Rows[0]["Calendar"].ToString() + "\n", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12)));
            centerParagraph.SpacingBefore = 10f;
            centerParagraph.SetLeading(15f, 0); // 50px space between lines
            centerParagraph.Add(new Chunk(model.dt2.Rows[0]["StartDate"].ToString() + " To " + model.dt2.Rows[0]["EndDate"].ToString() + "(" + model.dt2.Rows[0]["Venue"].ToString() + "- " + model.dt2.Rows[0]["OrganizatinShortName"].ToString() + ")", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)));
            centerCell.AddElement(centerParagraph);

            centerCell.HorizontalAlignment = Element.ALIGN_CENTER;
            centerCell.Padding = 8;
            titleTable.AddCell(centerCell);

            PdfPCell rightCell = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            {
                Border = Rectangle.TOP_BORDER | Rectangle.RIGHT_BORDER,
                Padding = 8
            };
            titleTable.AddCell(rightCell);

            doc.Add(titleTable);




            PdfPTable table = new PdfPTable(7) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 5f, 15f, 20f, 20f, 10f, 15f, 15f });

            // **Game/Unit Row**
            PdfPCell GameCell = new PdfPCell(new Phrase(model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " " + model.dt2.Rows[0]["OrganizatinShortName"].ToString(), FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, Font.ITALIC)))
            {
                Colspan = 12,
                HorizontalAlignment = Element.ALIGN_CENTER,
                Border = Rectangle.BOX,
                Padding = 5
            };
            table.AddCell(GameCell);


            string[] headers = { "Sr No.", "DESIGNATION", "NAME", "FATHER NAME", "PHOTO", "SIGNATURE", "CONTACT NO." };

            foreach (string header in headers)
                table.AddCell(CreateHeaderCell1(header));

            // **State/Unit Row**
            PdfPCell stateCell = new PdfPCell(new Phrase("State/Unit-" + model.OrganizatinShortName, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, Font.ITALIC)))
            {
                Colspan = 7,
                HorizontalAlignment = Element.ALIGN_LEFT,
                Border = Rectangle.BOX,
                Padding = 5
            };
            table.AddCell(stateCell);

            int i = 1;
            foreach (DataRow player in model.dt1.Rows)
            {
                table.AddCell(CreateBodyCell((i++).ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Type"]?.ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Name"]?.ToString(), Element.ALIGN_LEFT));
                table.AddCell(CreateBodyCell(player["FatherName"]?.ToString(), Element.ALIGN_LEFT));
                table.AddCell(GetPlayerImageCell(player["Photograph1"].ToString()));
                table.AddCell(CreateBodyCell("", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Phone"].ToString(), Element.ALIGN_CENTER));
            }

            doc.Add(table);
            doc.Add(new Paragraph("\n"));
        }
        private void AttendanceContent(Document doc, Players model)
        {
            PdfPTable sign = new PdfPTable(2) { WidthPercentage = 100 };
            sign.SetWidths(new float[] { 30f, 70f });
            sign.AddCell(CreateCell("Note:Mark ABST with Red Pen for Absent Candidate ", Element.ALIGN_LEFT, fontsize: 10));
            sign.AddCell(CreateCell("Declaration :- The attendence of above players have been physically taken by me & found correct according to above photographs placed before\r\nplayer's names.\n\n", Element.ALIGN_CENTER, fontsize: 10));
            doc.Add(sign);


            PdfPTable headerTable = new PdfPTable(3) { WidthPercentage = 100 };
            headerTable.SetWidths(new float[] { 46f, 27f, 27f });

            headerTable.AddCell(CreateCell("Coach/ Manager/ Cheif de mission", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Authorised officer deputed by Org. Secretary", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Organising Secretary\n", Element.ALIGN_LEFT));

            headerTable.AddCell(CreateCell("Sign...................................................\nReceived above players Participant Certificate...............................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Sign........................................................... ", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Sign................................................................. ", Element.ALIGN_LEFT, fontsize: 10));

            headerTable.AddCell(CreateCell("Name..............................................Name..........................................................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Name......................................................... ", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Name......................................................... ", Element.ALIGN_LEFT, fontsize: 10));

            headerTable.AddCell(CreateCell("Original Post.........................................Original Post......................................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Org.Post.........................................................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Org.Post.........................................................", Element.ALIGN_LEFT, fontsize: 10));

            headerTable.AddCell(CreateCell("Dept.Address....................................Dept.Address................................. ", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Dept.Address.................................................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Dept.Address.................................................", Element.ALIGN_LEFT, fontsize: 10));

            headerTable.AddCell(CreateCell("Mob No................................................Mob No..............................................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Mob No.................................................................", Element.ALIGN_LEFT, fontsize: 10));
            headerTable.AddCell(CreateCell("Mob No.................................................................", Element.ALIGN_LEFT, fontsize: 10));
            doc.Add(headerTable);

        }

        #endregion

        #region ID CARD

        public ActionResult AllIDCARDPDF(int Evt_Id)
        {
            Players model = new Players();
            GENERATE_DOWNLOAD_History model1 = new GENERATE_DOWNLOAD_History();
            model.dt2 = DbLayer.DisplayLists(24, Evt_Id);
            model.dt3 = DbLayer.DisplayLists(44, EventId: Evt_Id);

            model1.Action = 2;
            model1.Event_Id = Evt_Id;
            model1.Doc_Type = "AllIdCardSheetPDF";
            model1.Doc_Status = "AllIDCARDGenerate";
            model1.CreatedBy = int.Parse(Session["UserId"].ToString());
            model.dt4 = DbLayer.Attendance_Gen_History(model1);
            if (model.dt4.Rows.Count == 0)
            {
                using (MemoryStream memStream = new MemoryStream())
                {
                    Document doc = new Document(PageSize.A4.Rotate());
                    try
                    {
                        if (model.dt3.Rows.Count > 0)
                        {
                            string footer = model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString();
                            PdfWriter writer = PdfWriter.GetInstance(doc, memStream);
                            doc.Open();
                            PdfWatermarkHelper eventHandler = new PdfWatermarkHelper();
                            writer.PageEvent = eventHandler;
                            writer.PageEvent = new PDFFooter(footer);
                            foreach (DataRow item in model.dt3.Rows)
                            {
                                var data11 = dbHelper.ExecAdaptorDataTable("select OrganizatinName,OrganizatinShortName from M_Organizatins where IsActive=1 and OrganizatinId=" + Convert.ToInt32(item["Org_Id"]));
                                if (data11.Rows.Count > 0)
                                {
                                    model.OrganizatinShortName = data11.Rows[0]["OrganizatinShortName"].ToString();
                                    model.OrganizatinName = data11.Rows[0]["OrganizatinName"].ToString();
                                }
                                model.dt = DbLayer.DisplayLists(43, Convert.ToInt32(item["Org_Id"]), Evt_Id);
                                model.dt1 = DbLayer.DisplayLists(5, Convert.ToInt32(item["Org_Id"]), Evt_Id);
                                model.dt5 = DbLayer.DisplayLists(57, Convert.ToInt32(item["Org_Id"]), Evt_Id);

                                PlayersIDCardPDF(doc, model);
                                doc.NewPage();
                                OfficalIDCardPDF(doc, model);
                                doc.Add(new Paragraph("\n\n\n\n"));
                            }
                        }
                        doc.Close();

                        byte[] bytes = memStream.ToArray();
                        string fileName = DateTime.Now.Ticks + "_AllIdCardSheetPDF.pdf";
                        string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), fileName);
                        System.IO.File.WriteAllBytes(fullPath, bytes);


                        model1.Action = 1;
                        model1.Doc_Status = "AllIDCARDGenerate";
                        model1.DocumentName = fileName;
                        model.dt4 = DbLayer.Attendance_Gen_History(model1);


                        return File(bytes, "application/pdf", fileName);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                        return new HttpStatusCodeResult(500, "Internal Server Error");
                    }
                }
            }
            else
            {
                model1.Action = 1;
                model1.Doc_Status = "AllIDCARDPDF";
                model1.DocumentName = model.dt4.Rows[0][1].ToString();
                model.dt4 = DbLayer.Attendance_Gen_History(model1);
                string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), model1.DocumentName);
                return File(fullPath, "application/pdf", model1.DocumentName);
            }
        }

        public ActionResult IDCARDPDF(int Evt_Id, int OrgId)
        {
            Players model = new Players();
            model.EventId = Evt_Id;
            model.OrgId = OrgId;

            var data11 = dbHelper.ExecAdaptorDataTable("select OrganizatinName,OrganizatinShortName from M_Organizatins where IsActive=1 and OrganizatinId=" + OrgId);
            if (data11.Rows.Count > 0)
            {
                model.OrganizatinShortName = data11.Rows[0]["OrganizatinShortName"].ToString();
                model.OrganizatinName = data11.Rows[0]["OrganizatinName"].ToString();
            }

            model.dt = DbLayer.DisplayLists(43, OrgId, Evt_Id);
            model.dt1 = DbLayer.DisplayLists(5, OrgId, Evt_Id);
            model.dt5 = DbLayer.DisplayLists(57, OrgId, Evt_Id);
            model.dt2 = DbLayer.DisplayLists(24, Evt_Id);
            return GenerateIDCARDPDF(model);
        }
        private ActionResult GenerateIDCARDPDF(Players model)
        {
            GENERATE_DOWNLOAD_History model1 = new GENERATE_DOWNLOAD_History();

            model1.Action = 3;
            model1.Event_Id = model.EventId;
            model1.Org_Id = model.OrgId;
            model1.Doc_Type = "IdCardSheetPDF";
            model1.Doc_Status = "GenerateIDCARD";
            model1.CreatedBy = int.Parse(Session["UserId"].ToString());
            model.dt4 = DbLayer.Attendance_Gen_History(model1);
            if (model.dt4.Rows.Count == 0)
            {
                using (MemoryStream memStream = new MemoryStream())
                {
                    Document doc = new Document(PageSize.A4.Rotate());
                    try
                    {
                        string footer = model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString();
                        PdfWriter writer = PdfWriter.GetInstance(doc, memStream);
                        doc.Open();
                        PdfWatermarkHelper eventHandler = new PdfWatermarkHelper();
                        writer.PageEvent = eventHandler;
                        writer.PageEvent = new PDFFooter(footer);
                        PlayersIDCardPDF(doc, model);

                        doc.NewPage();

                        OfficalIDCardPDF(doc, model);

                        doc.Close();

                        byte[] bytes = memStream.ToArray();
                        string fileName = DateTime.Now.Ticks + "_IDCARDPDF.pdf";
                        string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), fileName);
                        System.IO.File.WriteAllBytes(fullPath, bytes);


                        model1.Action = 1;
                        model1.DocumentName = fileName;
                        model.dt4 = DbLayer.Attendance_Gen_History(model1);


                        return File(bytes, "application/pdf", fileName);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                        return new HttpStatusCodeResult(500, "Internal Server Error");
                    }
                }
            }
            else
            {
                model1.Action = 1;
                model1.Doc_Status = "DownloadIDCARD";
                model1.DocumentName = model.dt4.Rows[0][1].ToString();
                model.dt4 = DbLayer.Attendance_Gen_History(model1);
                string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), model1.DocumentName);
                return File(fullPath, "application/pdf", model1.DocumentName);
            }
        }

        public void PlayersIDCardPDF(Document doc, Players model)
        {
            PdfPTable mainTable = new PdfPTable(3) { WidthPercentage = 100 };
            mainTable.SetWidths(new float[] { 33f, 33f, 33f });

            Certificate Certi = new Certificate();
            Certi.Action = 1;
            Certi.CertificateTypeId = 1;
            Certi.CertificateIssuedBy = Convert.ToInt32(Session["UserId"]);
            foreach (DataRow item in model.dt.Rows)
            {
                Certi.CertificatePlayerId = item["PlayerId"].ToString();
                Certi.CertificateTeamRankId = model.dt5.Rows.Count > 0 ? Convert.ToInt32(model.dt5.Rows[0]["TeamRank_Id"]) : 0;
                var data = DbLayer.CertificateManage(Certi);

                PdfPTable idCardHeader = new PdfPTable(2) { WidthPercentage = 100 };
                idCardHeader.SetWidths(new float[] { 75f, 25f });

                Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.BLACK);
                Font monoFont = FontFactory.GetFont(FontFactory.COURIER, 10, BaseColor.BLACK);


                Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7);
                Font normalFont1 = FontFactory.GetFont(FontFactory.HELVETICA, 8);
                Font normalFont = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8);
                Font underlinedFont = new Font(normalFont1);
                underlinedFont.SetStyle(Font.UNDERLINE);


                PdfPTable headerTable = new PdfPTable(4) { WidthPercentage = 100 };
                headerTable.SetWidths(new float[] { 25f, 25f, 25f, 25f }); // Equal column width

                PdfPCell titleCell = new PdfPCell(new Phrase("VOLLEYBALL FEDERATION OF INDIA", underlinedFont))
                {
                    Colspan = 3,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = Rectangle.NO_BORDER
                };
                headerTable.AddCell(titleCell);

                headerTable.AddCell(new PdfPCell(GetPlayerImageCell(item["PlayerPhotograph"].ToString()))
                {
                    Rowspan = 3, // Image ko multiple rows cover karne ke liye
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    Border = Rectangle.NO_BORDER
                });


                PdfPCell eventCell = new PdfPCell(new Phrase($"" +
                    $"{model.dt2.Rows[0]["Calendar"].ToString().ToString()}\n" +
                    $"{model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString()}\n" +
                    $"{model.dt2.Rows[0]["StartDate"].ToString() + " To " + model.dt2.Rows[0]["EndDate"].ToString() + "\n(" + model.dt2.Rows[0]["Venue"].ToString() + "- " + model.dt2.Rows[0]["OrganizatinShortName"].ToString() + ")"}" +
                    $"", boldFont))
                {
                    Colspan = 3,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = Rectangle.NO_BORDER,
                    PaddingTop = 5
                };
                headerTable.AddCell(eventCell);

                headerTable.AddCell(new PdfPCell(new Phrase("IDENTITY CARD", underlinedFont)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = Rectangle.NO_BORDER, Colspan = 2 });

                headerTable.AddCell(new PdfPCell(new Phrase("Reg.No.:" + item["SGFIRegNo"].ToString(), normalFont))
                {
                    Colspan = 2,
                    PaddingTop = 8,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    Border = Rectangle.NO_BORDER
                });




                StringBuilder detailsBuilder = new StringBuilder();
                detailsBuilder.AppendLine("EVENT:    " + model.dt2.Rows[0]["GameName"].ToString());
                detailsBuilder.AppendLine("UNIT:     " + model.OrganizatinShortName);
                detailsBuilder.AppendLine("NAME:     " + item["Name"].ToString());


                detailsBuilder.AppendLine(
                    "FATHER'S NAME:".PadRight(20) +
                    $"{item["FatherName"].ToString()}".PadRight(20) +
                    "Secretary General".PadLeft(25)
                );


                detailsBuilder.AppendLine($"DOB:    {Convert.ToDateTime(item["DateOfBirth"]).ToString("dd/MM/yyyy")}      CLASS: {item["PlayerClass"].ToString()}`       ADM NO: {item["PlayerId"].ToString()}");
                detailsBuilder.AppendLine("SCHOOL:    " + item["SchoolName"].ToString());

                PdfPCell playerDetailsCell = new PdfPCell(new Phrase(detailsBuilder.ToString(), normalFont))
                {
                    Colspan = 4,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    Border = Rectangle.NO_BORDER
                };
                headerTable.AddCell(playerDetailsCell);


                mainTable.AddCell(new PdfPCell(headerTable) { Padding = 10, Border = Rectangle.NO_BORDER });
            }

            int remainingCells = 3 - (model.dt.Rows.Count % 3);
            if (remainingCells != 3)
            {
                for (int i = 0; i < remainingCells; i++)
                {
                    mainTable.AddCell(new PdfPCell() { Border = Rectangle.NO_BORDER });
                }
            }
            doc.Add(mainTable);
        }

        public void OfficalIDCardPDF(Document doc, Players model)
        {
            PdfPTable mainTable = new PdfPTable(3) { WidthPercentage = 100 };
            mainTable.SetWidths(new float[] { 33f, 33f, 33f });

            foreach (DataRow item in model.dt1.Rows)
            {
                PdfPTable idCardHeader = new PdfPTable(2) { WidthPercentage = 100 };
                idCardHeader.SetWidths(new float[] { 75f, 25f });

                Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.BLACK);
                Font monoFont = FontFactory.GetFont(FontFactory.COURIER, 10, BaseColor.BLACK);


                Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7);
                Font normalFont1 = FontFactory.GetFont(FontFactory.HELVETICA, 8);
                Font normalFont = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 8);
                Font underlinedFont = new Font(normalFont1);
                underlinedFont.SetStyle(Font.UNDERLINE);



                PdfPTable headerTable = new PdfPTable(4) { WidthPercentage = 100 };
                headerTable.SetWidths(new float[] { 25f, 25f, 25f, 25f }); // Equal column width

                PdfPCell titleCell = new PdfPCell(new Phrase("VOLLEYBALL FEDERATION OF INDIA", underlinedFont))
                {
                    Colspan = 3,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = Rectangle.NO_BORDER
                };
                headerTable.AddCell(titleCell);

                headerTable.AddCell(new PdfPCell(GetPlayerImageCell(item["Photograph1"].ToString()))
                {
                    Rowspan = 3, // Image ko multiple rows cover karne ke liye
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    Border = Rectangle.NO_BORDER
                });

                PdfPCell eventCell = new PdfPCell(new Phrase($"" +
                    $"{model.dt2.Rows[0]["Calendar"].ToString().ToString()}\n" +
                    $"{model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString()}\n" +
                    $"{model.dt2.Rows[0]["StartDate"].ToString() + " To " + model.dt2.Rows[0]["EndDate"].ToString() + "\n(" + model.dt2.Rows[0]["Venue"].ToString() + "- " + model.dt2.Rows[0]["OrganizatinShortName"].ToString() + ")"}" +
                    $"", boldFont))
                {
                    Colspan = 3,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    Border = Rectangle.NO_BORDER
                };
                headerTable.AddCell(eventCell);

                headerTable.AddCell(new PdfPCell(new Phrase("IDENTITY CARD", underlinedFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Border = Rectangle.NO_BORDER, Colspan = 3 });
                //headerTable.AddCell(new PdfPCell(new Phrase("")) {Border = Rectangle.NO_BORDER, Colspan = 1 });

                headerTable.AddCell(new PdfPCell(new Phrase(item["Type"].ToString(), normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Border = Rectangle.NO_BORDER, Colspan = 3 });
                headerTable.AddCell(new PdfPCell(new Phrase("")) { Border = Rectangle.NO_BORDER, Colspan = 1 });






                StringBuilder detailsBuilder = new StringBuilder();
                detailsBuilder.AppendLine("UNIT:     " + model.OrganizatinShortName.PadRight(50) + "Secretary General");
                detailsBuilder.AppendLine("NAME:     " + item["Name"].ToString());


                detailsBuilder.AppendLine(
                    "SO/WO:".PadRight(20) +
                    $"{item["FatherName"].ToString()}".PadRight(20)
                );



                PdfPCell playerDetailsCell = new PdfPCell(new Phrase(detailsBuilder.ToString(), normalFont))
                {
                    Colspan = 4,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    Border = Rectangle.NO_BORDER
                };
                headerTable.AddCell(playerDetailsCell);

                mainTable.AddCell(new PdfPCell(headerTable) { Padding = 10, Border = Rectangle.NO_BORDER });
            }

            int remainingCells = 3 - (model.dt1.Rows.Count % 3);
            if (remainingCells != 3)
            {
                for (int i = 0; i < remainingCells; i++)
                {
                    mainTable.AddCell(new PdfPCell() { Border = Rectangle.NO_BORDER });
                }
            }

            doc.Add(mainTable);
        }
        #endregion

        #region GamewiseStaticsReport
        public ActionResult GameWiseStatisticsPDF(int Action, int GameId, int AgeGroupId, string Gender, int C_Session = 0)
        {
            C_Session = C_Session == 0 ? Convert.ToInt32(Session["SessionId"]) : C_Session;
            Players player = new Players();
            string GenderGirl = null, GenderMale = null;
            switch (Gender)
            {
                case "BOYS":
                    GenderMale = "BOYS";
                    break;
                case "GIRLS":
                    GenderGirl = "GIRLS";
                    break;
                case "ALL":
                    GenderMale = "BOYS";
                    GenderGirl = "GIRLS";
                    break;
            }

            player.dt = DbLayer.GetPlayersCount(Action, GameId, AgeGroupId, GenderMale, GenderGirl, C_Session);
            player.dt1 = DbLayer.GetPlayersCount(Action + 1, GameId, AgeGroupId, GenderMale, GenderGirl, C_Session);

            using (MemoryStream memStream = new MemoryStream())
            {
                Document doc = new Document(PageSize.A4, 50, 50, 50, 50);
                PdfWriter.GetInstance(doc, memStream);
                doc.Open();

                GenerateStatisticsPDF(doc, player, Gender);

                doc.Close();

                return File(memStream.ToArray(), "application/pdf", "GamewiseStaticsReport.pdf");
            }
        }
        public void GenerateStatisticsPDF(Document doc, Players player, string Gender)
        {
            Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14, BaseColor.BLUE);
            Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.BLACK);
            Font contentFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, BaseColor.BLACK);
            Font footerFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, BaseColor.RED);

            // Subtitle (Event Details)
            Paragraph subTitle = new Paragraph(player.dt1.Rows[0]["Calendar"].ToString(), titleFont)
            {
                Alignment = Element.ALIGN_CENTER
            };
            doc.Add(subTitle);

            // Report Date
            Paragraph reportDate = new Paragraph("As on: " + DateTime.Today.ToString("dd-MM-yyyy"), contentFont)
            {
                Alignment = Element.ALIGN_CENTER
            };
            doc.Add(reportDate);

            // Event Information
            Paragraph eventInfo = new Paragraph(
                $"{player.dt1.Rows[0]["Calendar"]}\n{player.dt1.Rows[0]["StartDate"]} - {player.dt1.Rows[0]["EndDate"]}, " +
                $"{player.dt1.Rows[0]["Venue"]}, {player.dt1.Rows[0]["OrganizatinShortName"].ToString()}", contentFont)
            {
                Alignment = Element.ALIGN_CENTER
            };
            doc.Add(eventInfo);

            // Organizer Information
            Paragraph organizer = new Paragraph(
                $"Organised by SPORTS AUTHORITY OF {player.dt1.Rows[0]["Venue"]}, {player.dt1.Rows[0]["OrganizatinShortName"]}", contentFont)
            {
                Alignment = Element.ALIGN_CENTER
            };
            doc.Add(organizer);

            Paragraph organizer2 = new Paragraph("Under the aegis of Volleyball Federation of India\n\n", contentFont)
            {
                Alignment = Element.ALIGN_CENTER
            };
            doc.Add(organizer2);

            // Creating Table with 6 Columns
            PdfPTable table = new PdfPTable(5)
            {
                WidthPercentage = 100
            };
            table.SetWidths(new float[] { 1, 3, 1, 2, 1 });

            // Header Row
            AddCellToTable1(table, "S.No.", headerFont, 1, 3);
            AddCellToTable1(table, "STATE UNIT", headerFont, 1, 3);
            AddCellToTable(table, "ARCHERY", headerFont, 2, 1);
            AddCellToTable1(table, "Total", headerFont, 1, 3);

            // Sub-header Row
            AddCellToTable(table, "U-19", headerFont, 2, 1);
            if (Gender == "BOYS")
            {
                AddCellToTable(table, "Boys", headerFont, 2, 1);
            }
            else if (Gender == "GIRLS")
            {
                AddCellToTable(table, "Girls", headerFont, 2, 1);
            }
            else if (Gender == "ALL")
            {
                AddCellToTable(table, "Boys", headerFont, 1, 1);
                AddCellToTable(table, "Girls", headerFont, 1, 1);
            }

            int totalBoys = 0, totalGirls = 0, grandTotal = 0;
            int i = 1;
            foreach (DataRow item in player.dt.Rows)
            {
                int boys = int.Parse(item["BoysCount"].ToString());
                int girls = int.Parse(item["GirlsCount"].ToString());
                int total = int.Parse(item["TotalPlayers"].ToString());
                if (total == 0)
                    continue;

                table.AddCell(new PdfPCell(new Phrase(i.ToString(), contentFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });
                table.AddCell(new PdfPCell(new Phrase(item["OrganizatinShortName"].ToString(), contentFont)) { HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5 });

                if (Gender == "BOYS")
                {
                    table.AddCell(new PdfPCell(new Phrase(boys.ToString(), contentFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5, Colspan = 2 });
                }
                else if (Gender == "GIRLS")
                {
                    table.AddCell(new PdfPCell(new Phrase(girls.ToString(), contentFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5, Colspan = 2 });
                }
                else if (Gender == "ALL")
                {
                    table.AddCell(new PdfPCell(new Phrase(boys.ToString(), contentFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });
                    table.AddCell(new PdfPCell(new Phrase(girls.ToString(), contentFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });
                }

                table.AddCell(new PdfPCell(new Phrase(total.ToString(), contentFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });

                totalBoys += boys;
                totalGirls += girls;
                grandTotal += total;
            }

            // Table Footer - Total Row
            table.AddCell(new PdfPCell(new Phrase("", contentFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });
            table.AddCell(new PdfPCell(new Phrase("Total", footerFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });

            if (Gender == "BOYS")
            {
                table.AddCell(new PdfPCell(new Phrase(totalBoys.ToString(), footerFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5, Colspan = 2 });
            }
            else if (Gender == "GIRLS")
            {
                table.AddCell(new PdfPCell(new Phrase(totalGirls.ToString(), footerFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5, Colspan = 2 });
            }
            else if (Gender == "ALL")
            {
                table.AddCell(new PdfPCell(new Phrase(totalBoys.ToString(), footerFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });
                table.AddCell(new PdfPCell(new Phrase(totalGirls.ToString(), footerFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });
            }

            table.AddCell(new PdfPCell(new Phrase(grandTotal.ToString(), footerFont)) { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });

            doc.Add(table);

            // Footer Section
            Paragraph totalPlayers = new Paragraph($"\nTotal: {totalBoys} Boys, {totalGirls} Girls, {grandTotal} Players", headerFont)
            {
                Alignment = Element.ALIGN_CENTER
            };
            doc.Add(totalPlayers);

            Paragraph footer = new Paragraph("\nSGFI - PS To President", contentFont)
            {
                Alignment = Element.ALIGN_CENTER
            };
            doc.Add(footer);
        }
        static void AddCellToTable(PdfPTable table, string text, Font font, int colspan, int rowspan)
        {
            PdfPCell cell = new PdfPCell(new Phrase(text, font))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                Padding = 5,
                Colspan = colspan,
                Rowspan = rowspan
            };
            table.AddCell(cell);
        }
        static void AddCellToTable1(PdfPTable table, string text, Font font, int colspan, int rowspan)
        {
            PdfPCell cell = new PdfPCell(new Phrase(text, font))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                VerticalAlignment = Element.ALIGN_MIDDLE,
                Padding = 5,
                Colspan = colspan,
                Rowspan = rowspan
            };
            table.AddCell(cell);
        }
                
        #endregion

        #region DownloadMeritAttendence_INDIVIDUAL
        public ActionResult AttendanceSheetINDIVIDUALPDF()
        {
            Players model = new Players();
            /*var data11 = dbHelper.ExecAdaptorDataTable("select OrganizatinName,OrganizatinShortName from M_Organizatins where IsActive=1 and OrganizatinId=" + OrgId);
            if (data11.Rows.Count > 0)
            {
                model.OrganizatinShortName = data11.Rows[0]["OrganizatinShortName"].ToString();
                model.OrganizatinName = data11.Rows[0]["OrganizatinName"].ToString();
            }

            model.dt = DbLayer.DisplayLists(43, OrgId, Evt_Id);
            model.dt1 = DbLayer.DisplayLists(5, OrgId, Evt_Id);
            model.dt2 = DbLayer.DisplayLists(24, Evt_Id);*/
            return GenerateAttendanceSheetINDIVIDUALPDF(model);
        }
        private ActionResult GenerateAttendanceSheetINDIVIDUALPDF(Players model)
        {
            using (MemoryStream memStream = new MemoryStream())
            {
                Document doc = new Document(PageSize.A4.Rotate(), 20f, 20f, 20f, 20f);
                try
                {
                    PdfWriter writer = PdfWriter.GetInstance(doc, memStream);
                    doc.Open();

                    PlayersAttendanceINDIVIDUAL(doc, model);
                    doc.Close();
                    return File(memStream.ToArray(), "application/pdf", "AttendanceSheet.pdf");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    return new HttpStatusCodeResult(500, "Internal Server Error");
                }
            }
        }
        private void PlayersAttendanceINDIVIDUAL(Document doc, Players model)
        {
            PdfPTable titleTable = new PdfPTable(1) { WidthPercentage = 100 };
            titleTable.SetWidths(new float[] { 100f });
            PdfPCell HeaderleftCell = new PdfPCell(new Phrase("VOLLEYBALL FEDERATION OF INDIA", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9, BaseColor.WHITE)))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(163, 67, 123),
                Border = Rectangle.TOP_BORDER | Rectangle.LEFT_BORDER | Rectangle.BOTTOM_BORDER | Rectangle.RIGHT_BORDER,
                Padding = 8
            };
            titleTable.AddCell(HeaderleftCell);
            PdfPCell HeaderleftCell1 = new PdfPCell(new Phrase(" ATHLETICS U-19 BOYS RANCHI-JHARKHAND", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9, BaseColor.WHITE)))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(163, 67, 123),
                Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER,
                Padding = 8
            };
            titleTable.AddCell(HeaderleftCell1);


            doc.Add(titleTable);



            PdfPTable table = new PdfPTable(11) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 5f, 10f, 20f, 20f, 12f, 7f, 10f, 23f, 17f, 15f, 15f });

            string[] headers = { "SR", "PLACE", "CANDIDATE NAME", "FATHER NAME", "DOB", "CLASS", "UNIT", " MERIT NO", "CRT.NO.", "MOBILE NO.", "SIGNATURE" };

            foreach (string header in headers)
                table.AddCell(CreateHeaderCell1(header));

            PdfPCell stateCell = new PdfPCell(new Phrase("400 M HURDLE (U-19 BOYS) INDIVIDUAL", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, Font.ITALIC)))
            {
                Colspan = 11,
                HorizontalAlignment = Element.ALIGN_CENTER,
                Border = Rectangle.BOX,
                Padding = 5
            };
            table.AddCell(stateCell);

            int i = 1;

            /*foreach (DataRow player in model.dt.Rows)
            {*/
            for (int j = 0; j < 3; j++)
            {
                table.AddCell(CreateBodyCell((i++).ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell((i++).ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("BHUSHAN SUNIL PATIL", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("SUNIL", Element.ALIGN_CENTER));

                string dob = DateTime.TryParse("01-07-2006".ToString(), out DateTime dateOfBirth)
                             ? dateOfBirth.ToString("dd-MM-yyyy")
                             : "N/A";
                table.AddCell(CreateBodyCell("01-07-2006", Element.ALIGN_CENTER));

                table.AddCell(CreateBodyCell("12", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("KARNATAKA", Element.ALIGN_LEFT));
                table.AddCell(CreateBodyCell("HIGH PRIORITY-2024-25 RANCHI-JHARKHAND ATHLETICS(400 M HURDLE)U-19 BOYS 120650", Element.ALIGN_CENTER));

                //table.AddCell(GetPlayerImageCell(player["PlayerPhotograph"]?.ToString()));
                table.AddCell(CreateBodyCell("317-11444-120650", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("", Element.ALIGN_CENTER));
            }

            doc.Add(table);
            doc.Add(new Paragraph("\n"));
        }


        public ActionResult AttendanceSheetINDIVIDUAL_EXCEL()
        {
            Players model = new Players();
            /*var data11 = dbHelper.ExecAdaptorDataTable("select OrganizatinName,OrganizatinShortName from M_Organizatins where IsActive=1 and OrganizatinId=" + OrgId);
            if (data11.Rows.Count > 0)
            {
                model.OrganizatinShortName = data11.Rows[0]["OrganizatinShortName"].ToString();
                model.OrganizatinName = data11.Rows[0]["OrganizatinName"].ToString();
            }

            model.dt = DbLayer.DisplayLists(43, OrgId, Evt_Id);
            model.dt1 = DbLayer.DisplayLists(5, OrgId, Evt_Id);
            model.dt2 = DbLayer.DisplayLists(24, Evt_Id);*/
            return GenerateAttendanceSheetINDIVIDUAL_EXCEL(model);
        }
        public ActionResult GenerateAttendanceSheetINDIVIDUAL_EXCEL(Players model)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("AttendanceSheet");
                PlayersAttendanceINDIVIDUAL_EXCEL(worksheet, model);

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                using (MemoryStream stream = new MemoryStream())
                {
                    package.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "AttendanceSheet.xlsx");
                }
            }
        }

        private void PlayersAttendanceINDIVIDUAL_EXCEL(ExcelWorksheet worksheet, Players model)
        {
            worksheet.Cells[1, 1, 1, 11].Merge = true;
            worksheet.Row(1).Height = 30;
            worksheet.Cells[1, 1].Value = "VOLLEYBALL FEDERATION OF INDIA";
            worksheet.Cells[1, 1].Style.Font.Bold = true;
            worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(163, 67, 123));
            worksheet.Cells[1, 1].Style.Font.Color.SetColor(Color.White);

            worksheet.Cells[2, 1, 2, 11].Merge = true;
            worksheet.Row(2).Height = 30;
            worksheet.Cells[2, 1].Value = "ATHLETICS U-19 BOYS RANCHI-JHARKHAND";
            worksheet.Cells[2, 1].Style.Font.Bold = true;
            worksheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(163, 67, 123));
            worksheet.Cells[2, 1].Style.Font.Color.SetColor(Color.White);



            string[] headers = { "SR", "PLACE", "CANDIDATE NAME", "FATHER NAME", "DOB", "CLASS", "UNIT", "MERIT NO", "CRT.NO.", "MOBILE NO.", "SIGNATURE" };
            for (int col = 0; col < headers.Length; col++)
            {
                worksheet.Cells[3, col + 1].Value = headers[col];
                worksheet.Cells[3, col + 1].Style.Font.Bold = true;
                worksheet.Cells[3, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[3, col + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }


            worksheet.Cells[4, 1, 4, 11].Merge = true;
            worksheet.Cells[4, 1].Value = "400 M HURDLE (U-19 BOYS) INDIVIDUAL";
            worksheet.Cells[4, 1].Style.Font.Bold = true;
            worksheet.Cells[4, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[4, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);


            int row = 5;
            for (int i = 1; i <= 3; i++)
            {
                worksheet.Cells[row, 1].Value = i;
                worksheet.Cells[row, 2].Value = "PLACE";
                worksheet.Cells[row, 3].Value = "BHUSHAN SUNIL PATIL";
                worksheet.Cells[row, 4].Value = "SUNIL";
                worksheet.Cells[row, 5].Value = "01-07-2006";
                worksheet.Cells[row, 6].Value = "12";
                worksheet.Cells[row, 7].Value = "KARNATAKA";
                worksheet.Cells[row, 8].Value = "HIGH PRIORITY-2024-25 RANCHI-JHARKHAND";
                worksheet.Cells[row, 9].Value = "317-11444-120650";
                worksheet.Cells[row, 10].Value = "";
                worksheet.Cells[row, 11].Value = "";

                for (int col = 1; col <= 11; col++)
                {
                    worksheet.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                row++;
            }
        }
        #endregion

        #region DownloadMeritAttendence_TEAM
        public ActionResult AttendanceSheetTEAMPDF(int Evt_Id)
        {
            Players model = new Players();
            GENERATE_DOWNLOAD_History model1 = new GENERATE_DOWNLOAD_History();

            model1.Action = 3;
            model1.Event_Id = model.EventId;
            model1.Org_Id = model.OrgId;
            model1.Doc_Type = "AttendanceSheetTEAMPDF";
            model1.Doc_Status = "GenerateTEAM";
            model1.CreatedBy = int.Parse(Session["UserId"].ToString());
            model.dt4 = DbLayer.Attendance_Gen_History(model1);
            if (model.dt4.Rows.Count == 0)
            {
                model.dt2 = DbLayer.DisplayLists(24, Evt_Id);
                using (MemoryStream memStream = new MemoryStream())
                {
                    Document doc = new Document(PageSize.A4.Rotate(), 20f, 20f, 20f, 20f);
                    try
                    {
                        PdfWriter writer = PdfWriter.GetInstance(doc, memStream);
                        doc.Open();

                        model.dt3 = DbLayer.DisplayLists(44, EventId: Evt_Id);
                        PlayersAttendanceTEAMHeader(doc, model);
                        if (model.dt3.Rows.Count > 0)
                        {
                            string footer = model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString();
                            PdfWatermarkHelper eventHandler = new PdfWatermarkHelper();
                            writer.PageEvent = eventHandler;
                            writer.PageEvent = new PDFFooter(footer);
                            foreach (DataRow item in model.dt3.Rows)
                            {
                                var data11 = dbHelper.ExecAdaptorDataTable("select OrganizatinName,OrganizatinShortName from M_Organizatins where IsActive=1 and OrganizatinId=" + Convert.ToInt32(item["Org_Id"]));
                                if (data11.Rows.Count > 0)
                                {
                                    model.OrganizatinShortName = data11.Rows[0]["OrganizatinShortName"].ToString();
                                    model.OrganizatinName = data11.Rows[0]["OrganizatinName"].ToString();
                                }
                                model.dt = DbLayer.DisplayLists(43, Convert.ToInt32(item["Org_Id"]), Evt_Id);
                                model.dt5 = DbLayer.DisplayLists(57, Convert.ToInt32(item["Org_Id"]), Evt_Id);

                                PlayersAttendanceTEAM(doc, model);

                            }
                        }


                        doc.Close();

                        byte[] bytes = memStream.ToArray();
                        string fileName = DateTime.Now.Ticks + "_AttendanceSheetTEAM.pdf";
                        string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), fileName);
                        System.IO.File.WriteAllBytes(fullPath, bytes);

                        model1.Action = 1;
                        model1.DocumentName = fileName;
                        model.dt4 = DbLayer.Attendance_Gen_History(model1);

                        return File(bytes, "application/pdf", fileName);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                        return new HttpStatusCodeResult(500, "Internal Server Error");
                    }
                }
            }
            else
            {
                model1.Action = 1;
                model1.Doc_Type = "AttendanceSheetTEAMPDF";
                model1.Doc_Status = "DownloadTEAM";
                model1.DocumentName = model.dt4.Rows[0][1].ToString();
                model.dt4 = DbLayer.Attendance_Gen_History(model1);
                string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), model1.DocumentName);
                return File(fullPath, "application/pdf", model1.DocumentName);
            }

        }
        private void PlayersAttendanceTEAMHeader(Document doc, Players model)
        {

            PdfPTable titleTable = new PdfPTable(1) { WidthPercentage = 100 };
            titleTable.SetWidths(new float[] { 100f });
            PdfPCell HeaderleftCell = new PdfPCell(new Phrase("VOLLEYBALL FEDERATION OF INDIA", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9, BaseColor.WHITE)))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(163, 67, 123),
                Border = Rectangle.TOP_BORDER | Rectangle.LEFT_BORDER | Rectangle.BOTTOM_BORDER | Rectangle.RIGHT_BORDER,
                Padding = 8
            };
            titleTable.AddCell(HeaderleftCell);
            string Games = model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " " + model.dt2.Rows[0]["Venue"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString();

            PdfPCell HeaderleftCell1 = new PdfPCell(new Phrase(Games, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9, BaseColor.WHITE)))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(163, 67, 123),
                Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER,
                Padding = 8
            };
            titleTable.AddCell(HeaderleftCell1);


            doc.Add(titleTable);



            PdfPTable table = new PdfPTable(11) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 5f, 10f, 18f, 18f, 12f, 7f, 14f, 23f, 17f, 15f, 15f });

            string[] headers = { "SR", "PLACE", "CANDIDATE NAME", "FATHER NAME", "DOB", "CLASS", "UNIT", " MERIT NO", "CRT.NO.", "MOBILE NO.", "SIGNATURE" };

            foreach (string header in headers)
                table.AddCell(CreateHeaderCell1(header));
            doc.Add(table);
        }
        private void PlayersAttendanceTEAM(Document doc, Players model)
        {
            int i = 1;


            PdfPTable table = new PdfPTable(11) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 5f, 10f, 18f, 18f, 12f, 7f, 14f, 23f, 17f, 15f, 15f });

            Certificate Certi = new Certificate();
            Certi.Action = 1;
            Certi.CertificateTypeId = 1;
            Certi.CertificateIssuedBy = Convert.ToInt32(Session["UserId"]);
            foreach (DataRow player in model.dt.Rows)
            {
                Certi.CertificatePlayerId = player["PlayerId"].ToString();
                Certi.CertificateTeamRankId = model.dt5.Rows.Count > 0 ? Convert.ToInt32(model.dt5.Rows[0]["TeamRank_Id"]) : 0;
                var data = DbLayer.CertificateManage(Certi);

                table.AddCell(CreateBodyCell((i++).ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Rank_Id"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["Name"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["FatherName"].ToString(), Element.ALIGN_CENTER));

                string dob = DateTime.TryParse(player["DateOfBirth"].ToString(), out DateTime dateOfBirth)
                             ? dateOfBirth.ToString("dd-MM-yyyy")
                             : "N/A";
                table.AddCell(CreateBodyCell(dob, Element.ALIGN_CENTER));

                table.AddCell(CreateBodyCell(player["PlayerClass"].ToString(), Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell(player["OrganizatinShortName"].ToString(), Element.ALIGN_LEFT));
                string merit = $"HIGH PRIORITY {player["Session"].ToString()} {player["Venue"].ToString()}" +
                    $" {player["OrganizatinShortName"].ToString()} {player["GameName"].ToString()}" +
                    $" {player["AgeGroupName"].ToString()}";

                table.AddCell(CreateBodyCell(merit, Element.ALIGN_CENTER));

                //table.AddCell(GetPlayerImageCell(player["PlayerPhotograph"]?.ToString()));
                table.AddCell(CreateBodyCell(data.Rows.Count > 0 ? data.Rows[0]["Certificate_No"].ToString() : "", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("", Element.ALIGN_CENTER));
            }

            doc.Add(table);
        }



        public ActionResult AttendanceSheetTEAM_EXCEL(int Evt_Id)
        {
            Players model = new Players();
            GENERATE_DOWNLOAD_History model1 = new GENERATE_DOWNLOAD_History();
            model1.Action = 3;
            model1.Event_Id = model.EventId;
            model1.Org_Id = model.OrgId;
            model1.Doc_Type = "AttendanceSheetTEAM_EXCEL";
            model1.Doc_Status = "GenerateTEAMEXCEL";
            model1.CreatedBy = int.Parse(Session["UserId"].ToString());
            model.dt4 = DbLayer.Attendance_Gen_History(model1);
            if (model.dt4.Rows.Count == 0)
            {
                model.dt2 = DbLayer.DisplayLists(24, Evt_Id);
                model.dt3 = DbLayer.DisplayLists(44, EventId: Evt_Id);
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("AttendanceSheet");
                    PlayersAtteTEAM_EXCELHeader(worksheet, model);

                    if (model.dt3.Rows.Count > 0)
                    {
                        foreach (DataRow item in model.dt3.Rows)
                        {
                            var data11 = dbHelper.ExecAdaptorDataTable("select OrganizatinName,OrganizatinShortName from M_Organizatins where IsActive=1 and OrganizatinId=" + Convert.ToInt32(item["Org_Id"]));
                            if (data11.Rows.Count > 0)
                            {
                                model.OrganizatinShortName = data11.Rows[0]["OrganizatinShortName"].ToString();
                                model.OrganizatinName = data11.Rows[0]["OrganizatinName"].ToString();
                            }
                            model.dt = DbLayer.DisplayLists(43, Convert.ToInt32(item["Org_Id"]), Evt_Id);
                            model.dt5 = DbLayer.DisplayLists(57, Convert.ToInt32(item["Org_Id"]), Evt_Id);
                            PlayersAttendanceTEAM_EXCEL(worksheet, model);
                        }
                    }
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    using (MemoryStream stream = new MemoryStream())
                    {
                        byte[] bytes = stream.ToArray();
                        string fileName = DateTime.Now.Ticks + "_AttendanceTEAM_EXCEL.xlsx";
                        string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), fileName);
                        System.IO.File.WriteAllBytes(fullPath, bytes);

                        model1.Action = 1;
                        model1.DocumentName = fileName;
                        model.dt4 = DbLayer.Attendance_Gen_History(model1);

                        package.SaveAs(stream);
                        return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            else
            {
                model1.Action = 1;
                model1.Doc_Status = "DownloadTEAMEXCEL";
                model1.DocumentName = model.dt4.Rows[0][1].ToString();
                model.dt4 = DbLayer.Attendance_Gen_History(model1);
                string fullPath = Path.Combine(Server.MapPath("~/Media/AllDocument"), model1.DocumentName);

                return File(fullPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", model1.DocumentName);
            }
        }

        private void PlayersAtteTEAM_EXCELHeader(ExcelWorksheet worksheet, Players model)
        {

            worksheet.Cells[1, 1, 1, 11].Merge = true;
            worksheet.Row(1).Height = 30;
            worksheet.Cells[1, 1].Value = "VOLLEYBALL FEDERATION OF INDIA";
            worksheet.Cells[1, 1].Style.Font.Bold = true;
            worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(163, 67, 123));
            worksheet.Cells[1, 1].Style.Font.Color.SetColor(Color.White);

            string Games = model.dt2.Rows[0]["GameName"].ToString() + " " + model.dt2.Rows[0]["AgeGroupName"].ToString() + " " + model.dt2.Rows[0]["Gender"].ToString() + " " + model.dt2.Rows[0]["Venue"].ToString() + " - " + model.dt2.Rows[0]["OrganizatinShortName"].ToString();

            worksheet.Cells[2, 1, 2, 11].Merge = true;
            worksheet.Row(2).Height = 30;
            worksheet.Cells[2, 1].Value = Games;
            worksheet.Cells[2, 1].Style.Font.Bold = true;
            worksheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(163, 67, 123));
            worksheet.Cells[2, 1].Style.Font.Color.SetColor(Color.White);



            string[] headers = { "SR", "PLACE", "CANDIDATE NAME", "FATHER NAME", "DOB", "CLASS", "UNIT", "MERIT NO", "CRT.NO.", "MOBILE NO.", "SIGNATURE" };
            for (int col = 0; col < headers.Length; col++)
            {
                worksheet.Cells[3, col + 1].Value = headers[col];
                worksheet.Cells[3, col + 1].Style.Font.Bold = true;
                worksheet.Cells[3, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[3, col + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }


            /*worksheet.Cells[4, 1, 4, 11].Merge = true;
            worksheet.Cells[4, 1].Value = "400 M HURDLE (U-19 BOYS) INDIVIDUAL";
            worksheet.Cells[4, 1].Style.Font.Bold = true;
            worksheet.Cells[4, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[4, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);*/

        }
        private void PlayersAttendanceTEAM_EXCEL(ExcelWorksheet worksheet, Players model)
        {
            int row = 4, i = 1;
            Certificate Certi = new Certificate();
            Certi.Action = 1;
            Certi.CertificateTypeId = 1;
            Certi.CertificateIssuedBy = Convert.ToInt32(Session["UserId"]);
            foreach (DataRow player in model.dt.Rows)
            {
                Certi.CertificatePlayerId = player["PlayerId"].ToString();
                Certi.CertificateTeamRankId = model.dt5.Rows.Count > 0 ? Convert.ToInt32(model.dt5.Rows[0]["TeamRank_Id"]) : 0;
                var data = DbLayer.CertificateManage(Certi);

                string dob = DateTime.TryParse(player["DateOfBirth"].ToString(), out DateTime dateOfBirth)
                             ? dateOfBirth.ToString("dd-MM-yyyy")
                             : "N/A";
                string merit = $"HIGH PRIORITY {player["Session"].ToString()} {player["Venue"].ToString()}" +
                    $" {player["OrganizatinShortName"].ToString()} {player["GameName"].ToString()}" +
                    $" {player["AgeGroupName"].ToString()}";


                worksheet.Cells[row, 1].Value = i++;
                worksheet.Cells[row, 2].Value = player["Rank_Id"].ToString();
                worksheet.Cells[row, 3].Value = player["Name"].ToString();
                worksheet.Cells[row, 4].Value = player["FatherName"].ToString();
                worksheet.Cells[row, 5].Value = dob;
                worksheet.Cells[row, 6].Value = player["PlayerClass"].ToString();
                worksheet.Cells[row, 7].Value = player["OrganizatinShortName"].ToString();
                worksheet.Cells[row, 8].Value = merit;
                worksheet.Cells[row, 9].Value = data.Rows.Count > 0 ? data.Rows[0]["Certificate_No"].ToString() : "";
                worksheet.Cells[row, 10].Value = "";
                worksheet.Cells[row, 11].Value = "";

                for (int col = 1; col <= 11; col++)
                {
                    worksheet.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                row++;
            }
        }
        #endregion

        #region State Request        
        public ActionResult StateRequestsPDF()
        {
            Players model = new Players();
            /*var data11 = dbHelper.ExecAdaptorDataTable("select OrganizatinName,OrganizatinShortName from M_Organizatins where IsActive=1 and OrganizatinId=" + OrgId);
            if (data11.Rows.Count > 0)
            {
                model.OrganizatinShortName = data11.Rows[0]["OrganizatinShortName"].ToString();
                model.OrganizatinName = data11.Rows[0]["OrganizatinName"].ToString();
            }

            model.dt = DbLayer.DisplayLists(43, OrgId, Evt_Id);
            model.dt1 = DbLayer.DisplayLists(5, OrgId, Evt_Id);
            model.dt2 = DbLayer.DisplayLists(24, Evt_Id);*/
            return GenerateStateRequestsPDF(model);
        }
        private ActionResult GenerateStateRequestsPDF(Players model)
        {
            using (MemoryStream memStream = new MemoryStream())
            {
                Document doc = new Document(PageSize.A4, 30f, 30f, 130f, 80f);
                try
                {
                    PdfWriter writer = PdfWriter.GetInstance(doc, memStream);
                    writer.PageEvent = new StateRequestsHeader();
                    writer.PageEvent = new StateRequestsFooter();
                    doc.Open();

                    /*PdfContentByte canvas = writer.DirectContentUnder;
                    BaseColor bgColor = new BaseColor(250, 250, 250); // Light gray background
                    canvas.SetColorFill(bgColor);
                    canvas.Rectangle(0, 0, PageSize.A4.Width, PageSize.A4.Height);
                    canvas.Fill();*/

                    PlayersStateRequestsTitle(doc, model);
                    doc.Close();
                    return File(memStream.ToArray(), "application/pdf", "Guidelines13_VolleyBall.pdf");
                }
                catch (Exception ex)
                {
                    //Guidelines13_ARCHERY_U11_Boys20(1).pdf
                    Console.WriteLine("Error: " + ex.Message);
                    return new HttpStatusCodeResult(500, "Internal Server Error");
                }
            }
        }

        private void PlayersStateRequestsTitle(Document doc, Players model)
        {
            Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, Font.UNDERLINE);
            //Font titleFont2 = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
            Font titleFont3 = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, Font.UNDERLINE);
            Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
            Font normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 9.5f);
            Font normalFont2 = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9);
            Font normalFont3 = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9, Font.UNDERLINE);



            Paragraph title = new Paragraph("Guidelines to Conduct National School Games 2024-25", titleFont);
            title.Alignment = Element.ALIGN_CENTER;
            doc.Add(title);
            doc.Add(new Paragraph("\n"));



            PdfPTable titleTable = new PdfPTable(3) { WidthPercentage = 100 };
            titleTable.SetWidths(new float[] { 33f, 34f, 33f });
            PdfPCell leftContent = new PdfPCell(new Phrase("From.\n Director, Sports & Youth Welfare Bihar", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, Font.ITALIC)))
            {
                Colspan = 3,
                HorizontalAlignment = Element.ALIGN_LEFT,
                Border = Rectangle.NO_BORDER,
                Padding = 5
            };
            titleTable.AddCell(leftContent);
            PdfPCell SpaceContent = new PdfPCell(new Phrase("")) { Border = Rectangle.NO_BORDER };
            titleTable.AddCell(SpaceContent);
            titleTable.AddCell(SpaceContent);
            PdfPCell DateContent = new PdfPCell(new Phrase("Date 19/09/2024", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, Font.ITALIC)))
            {
                Colspan = 1,
                HorizontalAlignment = Element.ALIGN_LEFT,
                Border = Rectangle.NO_BORDER,
                Padding = 5
            };
            titleTable.AddCell(DateContent);
            doc.Add(titleTable);
            //doc.Add(new Paragraph("\n"));

            Paragraph recipient = new Paragraph();
            recipient.Add(new Chunk("To,\n", normalFont));
            recipient.Add(new Chunk("           The President\n", boldFont));
            recipient.Add(new Chunk("             School Games Federation of India\n", normalFont));
            recipient.Add(new Chunk("             Campus of the Directorate of Secondary Education\n", normalFont));
            recipient.Add(new Chunk("             18, Park Road, Lucknow (Uttar Pradesh)\n\n", normalFont));
            recipient.Add(new Chunk("Sub.:      Guidelines for 68th National School Games( ARCHERY U-11 Boys) Championship 2024-25 at( BIHAR).\n", boldFont));
            recipient.Add(new Chunk("Ref.:\n\n", boldFont));
            recipient.Add(new Chunk("Dear Sir,\n", normalFont));
            recipient.Add(new Chunk("               It is my  privilege to inform you that the ", normalFont));
            recipient.Add(new Chunk("It is my  privilege to inform you that the ", normalFont));
            recipient.Add(new Chunk(" Director, Sports & Youth Welfare                         Bihar, (Name of Department) ", boldFont));
            var Gender = new Chunk("Boys", boldFont);
            var State = new Chunk("BIHAR", boldFont);
            //It is my privilege to inform you that the It is my privilege to inform you that the Director, Sports & Youth Welfare
            recipient.Add(new Chunk($"has been entrusted with the responsiblity of hosting 68th National School Games                               " +
                $"ARCHERY 2024-25U-11{Gender}competition under the auspices of School Games Federation of india. On behalf of the                               " +
                $"organizing committee,I take this opportunity to extend invitation to the respective contingent of all the States/UTs/Units to                       " +
                $"participate in the meet at {State}. The following are the detais related to the conduct of tournament:\n", normalFont));

            recipient.Add(new Chunk("(1) National School Games Summary:-\n\n", titleFont));
            doc.Add(recipient);



            PdfPTable table = new PdfPTable(5) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 20f, 20f, 20f, 20f, 20f });
            string[] headers = { "State", " Name Of Place", "Discipline", "Age Group", "Date of Championship" };
            foreach (string header in headers)
                table.AddCell(CreateHeaderCell1(header));
            /*int i = 1;
            foreach (DataRow player in model.dt.Rows)
            {*/
            for (int i = 1; i <= 1; i++)
            {
                table.AddCell(CreateBodyCell("BIHAR", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("Director Sports & Youth Welfare Mohenul Haq Stadium Bihar", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("ARCHERY", Element.ALIGN_LEFT));
                table.AddCell(CreateBodyCell("U-11 (Boys)", Element.ALIGN_LEFT));

                Paragraph recipient1 = new Paragraph();
                recipient1.Add(new Chunk("From:", boldFont));
                recipient1.Add(new Chunk("19th to 30th September,2024", normalFont));
                recipient1.Add(new Chunk("Reporting Date:", boldFont));
                recipient1.Add(new Chunk("30th September, 2024", normalFont));
                recipient1.Add(new Chunk("Last Online Entry:", boldFont));
                recipient1.Add(new Chunk("18th September, 2024", normalFont));
                table.AddCell(recipient1);

            }
            doc.Add(table);

            Paragraph Weather_rec = new Paragraph();
            var Month = new Chunk("September");
            Weather_rec.Add(new Chunk("(2)    ", boldFont));
            Weather_rec.Add(new Chunk("Weather:-\n", titleFont3));
            Weather_rec.Add(new Chunk("         During the month of", normalFont));
            Weather_rec.Add(new Chunk($" {Month} ", boldFont));
            Weather_rec.Add(new Chunk("weather in", normalFont));
            Weather_rec.Add(new Chunk($" {State} ", boldFont));
            Weather_rec.Add(new Chunk("is .You are advised to bring enough/suitable clothing accordingly.\n", normalFont));
            doc.Add(Weather_rec);

            Weather_rec.Clear();
            Month = new Chunk("19/09/2024 up to 18:40.");
            Weather_rec.Add(new Chunk("(3)    ", boldFont));
            Weather_rec.Add(new Chunk("Reporting Date:-\n", titleFont3));
            Weather_rec.Add(new Chunk("         The contingent of your state is expected to reach", normalFont));
            Weather_rec.Add(new Chunk($" {Month}", boldFont));
            Weather_rec.Add(new Chunk(" your arrival, at the  the eligibility form of the State               teams  and physical verification of players will be scrutinized on the same date at 19/09/2024\n", normalFont));
            doc.Add(Weather_rec);

            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(4)    ", boldFont));
            Weather_rec.Add(new Chunk("Place of Reporting:-\n", titleFont3));
            Weather_rec.Add(new Chunk("          (i) Place & location of control Room                      : Lucknow\n", normalFont));
            Weather_rec.Add(new Chunk("          (ii) Name Of Control Room in-Charge                   : Shubham\n", normalFont));
            Weather_rec.Add(new Chunk("          (iii) Mobile Nos. of Control Room in-Charge          : 8786876786\n", normalFont));
            doc.Add(Weather_rec);

            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(5)    ", boldFont));
            Weather_rec.Add(new Chunk("Reception\n", titleFont3));
            Weather_rec.Add(new Chunk("           Arrangements have been made for your reception at Railway Station & Bus  Stand.\n\n", normalFont));
            doc.Add(Weather_rec);


            table.DeleteBodyRows();
            table = new PdfPTable(4) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 25f, 25f, 25f, 25f });
            table.AddCell(CreateHeaderCell1("Reception Venue"));
            table.AddCell(CreateHeaderCell1("Name of Receptionist"));
            table.AddCell(CreateHeaderCell1("Mob./Ph.No."));
            table.AddCell(CreateHeaderCell1("Place and Time of Reception Counter"));
            for (int i = 1; i <= 1; i++)
            {
                table.AddCell(CreateBodyCell("Lucknow", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("Kanha", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("8267908509", Element.ALIGN_LEFT));
                table.AddCell(CreateBodyCell("Lucknow", Element.ALIGN_LEFT));
            }
            doc.Add(table);

            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(6)    ", boldFont));
            Weather_rec.Add(new Chunk("Identity Card:-\n", titleFont3));
            Weather_rec.Add(new Chunk("         All players must have identity card  duly signed/sttested by the Head of Controlling Officer Compentent authority", normalFont));
            doc.Add(Weather_rec);

            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(7)    ", boldFont));
            Weather_rec.Add(new Chunk("Entry of Teams:- \n", titleFont3));
            Weather_rec.Add(new Chunk("         You are requested to forward the information regarding initial entry of your participation before 10 days starting of the tournament positively to organizer.", normalFont));
            doc.Add(Weather_rec);

            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(8)    ", boldFont));
            Weather_rec.Add(new Chunk("How to Reach:-\n", titleFont3));
            Weather_rec.Add(new Chunk("          (i) Nearest Railway Station / Bus Stand from Hostel/ Control room with distance 1\n", normalFont));
            Weather_rec.Add(new Chunk("          (ii) Name of the nearest Junction from the hostel/accommodation/control room with distance 009\n", normalFont));
            Weather_rec.Add(new Chunk("          (iii) Name of all Railway Station / junction whichever  ( if  available )  with Distance Lucknow\n", normalFont));
            doc.Add(Weather_rec);


            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(9)    ", boldFont));
            Weather_rec.Add(new Chunk("PROVISIONAL PROGRAMME DATE\n\n", titleFont3));
            doc.Add(Weather_rec);

            table.DeleteBodyRows();
            table.AddCell(CreateHeaderCell1("Date"));
            table.AddCell(CreateHeaderCell1("Time"));
            table.AddCell(CreateHeaderCell1("Programme"));
            table.AddCell(CreateHeaderCell1("Place"));
            for (int i = 1; i <= 2; i++)
            {
                table.AddCell(CreateBodyCell("19/09/2024", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("17:43", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("4", Element.ALIGN_LEFT));
                table.AddCell(CreateBodyCell("Ambedkar Nagar", Element.ALIGN_LEFT));
            }
            doc.Add(table);
            doc.Add(new Paragraph("\n"));


            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(10)    ", boldFont));
            Weather_rec.Add(new Chunk("TRANSPORTAION:-\n", titleFont3));
            Weather_rec.Add(new Chunk("          (i) Dropping of Players  by bus from railway station/bus Stand to  Participants' accommodation, the condition of the bus should be according to weather,make them seated on the bus and should not be overloaded.\n", normalFont));
            Weather_rec.Add(new Chunk("          (ii) Arrangements for taking coach / sporting staff to accommodation place by a small car.\n", normalFont));
            Weather_rec.Add(new Chunk("          (iii) Pick and drop arrangemnt for all delegations from accomodation place to tournament venue.\n", normalFont));
            Weather_rec.Add(new Chunk("          (iv) For Chief-de-mission /President /secrearty General a car / appropriate vehicle should be arranged during the entire NSG.\n", normalFont));
            Weather_rec.Add(new Chunk("          (v)The Travelling expenses fromhome to  Competition Venue and return is responsibility of participating teams.\n", normalFont));
            doc.Add(Weather_rec);


            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(11)    ", boldFont));
            Weather_rec.Add(new Chunk("FINANCE & INSURANCE:-\n", titleFont3));
            Weather_rec.Add(new Chunk($"         The Organizing Commitee is responsible for  participant's accommodation & transportation " +
                $"arrangements during tournment and all technical arrangements in connection with the event.Each affliated unit must" +
                $" have insurance for all members of its delegation. Including compulsory insurance cover particularly health,accident & " +
                $"travel insurance for all the member of its delegation by participation unit.During the travel or competition any accident of team" +
                $"members,School Games Federation of india shall not be responsible for any claim.\n\n", normalFont));
            doc.Add(Weather_rec);

            table.DeleteBodyRows();
            table.AddCell(CreateHeaderCell1("Name of the State having Stay"));
            table.AddCell(CreateHeaderCell1("Name of the venue School/Hotel"));
            table.AddCell(CreateHeaderCell1("Name of responsible person for accommodation"));
            table.AddCell(CreateHeaderCell1("Telephone/Mob.No."));
            for (int i = 1; i <= 2; i++)
            {
                table.AddCell(CreateBodyCell("31", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("Aviral hotel", Element.ALIGN_CENTER));
                table.AddCell(CreateBodyCell("Aviral", Element.ALIGN_LEFT));
                table.AddCell(CreateBodyCell("8787876867", Element.ALIGN_LEFT));
            }
            doc.Add(table);

            Weather_rec.Clear();
            Weather_rec.Add(new Chunk($"         (1) Accommodation of players room category 3 star / Equivalent accommodation for 3 players in one room.\n", normalFont));
            Weather_rec.Add(new Chunk($"         (2) Coaches/Managers /sporting staff NTO /ITO room category 3 star or equivalent accommodation for 2 players in one room.\n", normalFont));
            Weather_rec.Add(new Chunk($"         (3) Chirf-de-mission President/General Secretary of NSF from the state,sport related NDF room category 03 star or upgrade accommodation for 01 officer in a single room as per the rules of the organizer.\n", normalFont));
            doc.Add(Weather_rec);


            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(13)    ", boldFont));
            Weather_rec.Add(new Chunk("MESS ARRANGEMENTS:-\n", titleFont3));
            Weather_rec.Add(new Chunk("   (a)", normalFont2));
            Weather_rec.Add(new Chunk("Common Mess:-\n", normalFont3));
            Weather_rec.Add(new Chunk("          Food will be available from  the common mess on payment as per SGFI'snorms at Rs.250/- per head per day. Food will be provided from common mess from 16:43 onwords.", normalFont));
            Weather_rec.Add(new Chunk("\n          Name of place :-Lucknow", normalFont));
            Weather_rec.Add(new Chunk("\n          Food in Common Mess: Veg. & Non-Veg", normalFont));
            Weather_rec.Add(new Chunk("\n   (b)", normalFont2));
            Weather_rec.Add(new Chunk("Mess for Technical Officials / VIPs:-  ", normalFont3));
            Weather_rec.Add(new Chunk("There will be separate mess for the Technical Officials /VIPs.\n", normalFont));
            Weather_rec.Add(new Chunk("   (c)", normalFont2));
            Weather_rec.Add(new Chunk("Common mess menu:-  ", normalFont3));
            Weather_rec.Add(new Chunk("As per School Games Federation of India norms for the common mess menu is as follows:-\n", normalFont));
            Weather_rec.Add(new Chunk("         Meals(buffet style ) will be served three times a day, Meal times are as follows-\n\n", normalFont));

            doc.Add(Weather_rec);


            table.DeleteBodyRows();
            table = new PdfPTable(3) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 33f, 34f, 33f });

            table.AddCell(CreateBodyCell("Breakfast", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell(" :", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("19:44   to   20:44", Element.ALIGN_LEFT));

            table.AddCell(CreateBodyCell("Lunch", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell(" :", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("19:44   to   20:44", Element.ALIGN_LEFT));

            table.AddCell(CreateBodyCell("Dinner", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell(" :", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("19:44   to   20:44", Element.ALIGN_LEFT));

            doc.Add(table);

            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("   (d)", normalFont2));
            Weather_rec.Add(new Chunk("Common Mess Menu:- \n\n", normalFont3));
            doc.Add(Weather_rec);

            table.DeleteBodyRows();
            table = new PdfPTable(2) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 50f, 50f });

            table.AddCell(CreateBodyCell("Breakfast", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("19:44   to   20:44", Element.ALIGN_LEFT));

            table.AddCell(CreateBodyCell("Lunch", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("19:44   to   20:44", Element.ALIGN_LEFT));

            table.AddCell(CreateBodyCell("Dinner", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("19:44   to   20:44", Element.ALIGN_LEFT));
            doc.Add(table);


            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(14)    ", boldFont));
            Weather_rec.Add(new Chunk("Composition of Team:-\n", titleFont3));
            Weather_rec.Add(new Chunk($"         Each affiliated Unit/UT/State can send only one team in each category.Team will be consisting as follows:\n\n", normalFont));
            doc.Add(Weather_rec);

            table.DeleteBodyRows();
            table = new PdfPTable(7) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 10f, 15f, 25f, 15f, 15f, 15f, 5f });
            table.AddCell(CreateHeaderCell1("No."));
            table.AddCell(CreateHeaderCell1("Discipline"));
            table.AddCell(CreateHeaderCell1(""));
            table.AddCell(CreateHeaderCell1("Coach"));
            table.AddCell(CreateHeaderCell1("Manager"));
            table.AddCell(CreateHeaderCell1("Total"));
            table.AddCell(CreateHeaderCell1("1"));
            doc.Add(table);
            Weather_rec.Clear();
            Weather_rec.Add(new Chunk($"         Please note that for all Games there will be only One Chief-de-Mission from each State. Chief-de-Mission should be of the status of minimum Deputy Director.\n", normalFont));
            doc.Add(Weather_rec);


            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(15)    ", boldFont));
            Weather_rec.Add(new Chunk("Eligibility Criteria:-\n", titleFont3));
            Weather_rec.Add(new Chunk("   (a)", normalFont2));
            Weather_rec.Add(new Chunk("          Player(Boy or girl) a regular enrolled  student of school is classified under following categories:", normalFont));
            Weather_rec.Add(new Chunk("\n    1.UNDER11yrs- ", normalFont2));
            Weather_rec.Add(new Chunk("Player should be ", normalFont));
            Weather_rec.Add(new Chunk("Minimum 08 yrs and less than 11 yrs ", normalFont2));
            Weather_rec.Add(new Chunk("and studying ", normalFont));
            Weather_rec.Add(new Chunk("between Class-3 to Class 5. ", normalFont2));
            Weather_rec.Add(new Chunk("Student/player ", normalFont));
            Weather_rec.Add(new Chunk("below 3rd standard ", normalFont2));
            Weather_rec.Add(new Chunk("will not be eligible to participate in the SGFI NSGs games. ", normalFont));

            Weather_rec.Add(new Chunk("\n    2.UNDER 14,17& 19 years- ", normalFont2));
            Weather_rec.Add(new Chunk("Player/Student studying ", normalFont));
            Weather_rec.Add(new Chunk("below 6th standard ", normalFont2));
            Weather_rec.Add(new Chunk("will not be eligible to participate in the SGFI NDGs games.", normalFont));

            Weather_rec.Add(new Chunk("\n    3.", normalFont2));
            Weather_rec.Add(new Chunk("Any student / Player who has passed", normalFont));
            Weather_rec.Add(new Chunk("12th standard ", normalFont2));
            Weather_rec.Add(new Chunk("will ", normalFont));
            Weather_rec.Add(new Chunk("not be eligible to participate ", normalFont2));
            Weather_rec.Add(new Chunk("in the SGFI NSGs games irrespective of being in any age category. \n", normalFont));

            Weather_rec.Add(new Chunk("   (b) Eligibility/Age certificate :-\n", normalFont2));
            Weather_rec.Add(new Chunk("          it is mandatory for all players to have ", normalFont));
            Weather_rec.Add(new Chunk("AADHAAR No./10th class marksheet/Date of Birth Certificate(should be issued minimum 5 years before),", normalFont2));
            Weather_rec.Add(new Chunk("Official Entry & eligibility forms in new format duly signed/attested by the head of the institution/ principal & counter signature by the competent suthority of State/Unit/UT.The team manager will be responsible for bringing the eligibility/birth certificate of the participants,which are to be handed over to the organizing committee. in tournament only official entry form signed by the compentent authority of State/UT/Unit will be acceptable.In the lack of this signed official entry form ,it is not possible to participate in the tournament & issue the merit/participation certificate.\n\n", normalFont));

            doc.Add(Weather_rec);



            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(16)    ", boldFont));
            Weather_rec.Add(new Chunk("ANTI-DOPING CLINIC FOR PARTICIPANTS/OFFICIALS:\n\n", titleFont3));
            doc.Add(Weather_rec);

            table.DeleteBodyRows();
            table = new PdfPTable(4) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 25f, 25f, 25f, 25f });
            table.AddCell(CreateHeaderCell1("Place of Managers Meeting"));
            table.AddCell(CreateHeaderCell1("Date"));
            table.AddCell(CreateHeaderCell1("Time"));
            table.AddCell(CreateHeaderCell1("Name of Organzing In\r\ncharge with Mob.No."));

            table.AddCell(CreateBodyCell("Lucknow", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("11/09/2024", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell(" 17:47", Element.ALIGN_LEFT));
            table.AddCell(CreateBodyCell("Amit", Element.ALIGN_LEFT));

            doc.Add(table);


            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(17)    ", boldFont));
            Weather_rec.Add(new Chunk("CHIEF-DE-MISSION/H.O.D MEETING:-\n\n", titleFont3));
            doc.Add(Weather_rec);

            table.DeleteBodyRows();
            table = new PdfPTable(5) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 20f, 20f, 20f, 20f, 20f });
            table.AddCell(CreateHeaderCell1("Place of Managers Meeting"));
            table.AddCell(CreateHeaderCell1("Date"));
            table.AddCell(CreateHeaderCell1("Time"));
            table.AddCell(CreateHeaderCell1("Name of Organzing In-charge"));
            table.AddCell(CreateHeaderCell1("Mob.No."));

            table.AddCell(CreateBodyCell("Lucknow", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("11/09/2024", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell(" 17:47", Element.ALIGN_LEFT));
            table.AddCell(CreateBodyCell("Amit", Element.ALIGN_LEFT));
            table.AddCell(CreateBodyCell("8678687687", Element.ALIGN_LEFT));

            doc.Add(table);

            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(18)    ", boldFont));
            Weather_rec.Add(new Chunk("COACHES MEETING:-\n\n", titleFont3));
            doc.Add(Weather_rec);

            table.DeleteBodyRows();
            table = new PdfPTable(5) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 20f, 20f, 20f, 20f, 20f });
            table.AddCell(CreateHeaderCell1("Place of Managers Meeting"));
            table.AddCell(CreateHeaderCell1("Date"));
            table.AddCell(CreateHeaderCell1("Time"));
            table.AddCell(CreateHeaderCell1("Name of Organzing In-charge"));
            table.AddCell(CreateHeaderCell1("Mob.No."));

            table.AddCell(CreateBodyCell("Lucknow", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("18/09/2024", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell(" 17:48", Element.ALIGN_LEFT));
            table.AddCell(CreateBodyCell("Ram", Element.ALIGN_LEFT));
            table.AddCell(CreateBodyCell("8767687687", Element.ALIGN_LEFT));

            doc.Add(table);


            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(10)    ", boldFont));
            Weather_rec.Add(new Chunk("DOCUMENTS SUBMISSION :-\n", titleFont3));
            Weather_rec.Add(new Chunk("          (1)  General manager of all States/UTs/Units must bring & produce the AUTHORITY LETTER from their competent authority for attestation power / signature on Eligibility Certificates / Entry forms to the organizers / School Games Federation of India personnel's.\n", normalFont));
            Weather_rec.Add(new Chunk("          (2) The State Flag (of your state) of 6fl.x 4ft.size.       -       2\n", normalFont));
            Weather_rec.Add(new Chunk("          (3) Fuly Filled Eligibility Certificate         -   In Triplicate \n", normalFont));
            Weather_rec.Add(new Chunk("          (4) Copy of AADHAAR Card            - 01\n", normalFont));
            Weather_rec.Add(new Chunk("          (5) Complete list of participants & officials.           -   Original\n", normalFont));
            Weather_rec.Add(new Chunk("          (6) Certified that each one of above players is born on or after", normalFont));
            Weather_rec.Add(new Chunk(" ... Yrs", normalFont2));
            Weather_rec.Add(new Chunk("hence they are eligible for participating in their respective age group.This certificate will be issued only by competent authority of State/UT/Unit", normalFont));
            doc.Add(Weather_rec);



            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(20)    ", boldFont));
            Weather_rec.Add(new Chunk("Other details & further contact information  for Nodal Officer of National School Games:-\n\n", titleFont3));
            doc.Add(Weather_rec);

            table.DeleteBodyRows();
            table = new PdfPTable(4) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 25f, 25f, 25f, 25f });
            table.AddCell(CreateHeaderCell1("Name of Nodal Officer"));
            table.AddCell(CreateHeaderCell1("Designation & Correspondence address"));
            table.AddCell(CreateHeaderCell1("Ph./Mob.No./Fax No./E Mail Address"));
            table.AddCell(CreateHeaderCell1("Official Website of national School Games"));

            table.AddCell(CreateBodyCell("Arun", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("Lucknow", Element.ALIGN_CENTER));
            table.AddCell(CreateBodyCell("8768657685", Element.ALIGN_LEFT));
            table.AddCell(CreateBodyCell("www.www.com", Element.ALIGN_LEFT));
            doc.Add(table);





            Weather_rec.Clear();
            Weather_rec.Add(new Chunk("(21)   ", boldFont));
            Weather_rec.Add(new Chunk("Online Entry:-\n", titleFont3));
            Weather_rec.Add(new Chunk("          Before the start of online registration of players, prepare a Demand Draft (D.D.) of the total amouns favoring", normalFont));
            Weather_rec.Add(new Chunk(" \"School Games Federation of India\",", normalFont2));
            Weather_rec.Add(new Chunk(" payable at ,, of the total amount,for the team @Rs. 200/- per player for the team.Fill the", normalFont));
            Weather_rec.Add(new Chunk(" D.D. Details, after Verifying EVENT CODE and PASSWORD,", normalFont2));
            Weather_rec.Add(new Chunk(" to begin the Online Registration Process. The Original D.D. is to be deposited with the representative of School Games Fedreation of India, at the competition venue and receipt to be collected from them.\n", normalFont));

            Weather_rec.Add(new Chunk("(22)   ", boldFont));
            Weather_rec.Add(new Chunk("Media:-\n", titleFont3));
            Weather_rec.Add(new Chunk("          The entire work of print media and electronic media will be done by the organizing committee and it will be monitored by SGFI media cell. Minimum 2 pre-event, Press conferences should be done with the presence of SGFI office bearers and staff.\n", normalFont));
            Weather_rec.Add(new Chunk("NSG will be monitored from SGFI control room through electronic surveillance system.", normalFont));
            Weather_rec.Add(new Chunk(" \"School Games Federation of India\",", normalFont2));

            Weather_rec.Add(new Chunk("\n(23)   ", boldFont));
            Weather_rec.Add(new Chunk("Timing, Scoring & Result:-\n", titleFont3));
            Weather_rec.Add(new Chunk("          Timing, Scoring & Result system should be in place by the authorized vendor. Minimum 2 LED walls of 20x12 should be in place for LIVE scoring/Telecasting at the Venue, LIVE Telecast should also be in place with appropriate number of cameras and a live link should be provided to telecast on SGFI website and social media platforms.\n", normalFont));

            Weather_rec.Add(new Chunk("(24)   ", boldFont));
            Weather_rec.Add(new Chunk("Branding of Stadium & Host City:-\n", titleFont3));
            Weather_rec.Add(new Chunk("          Proper branding of stadium should be in place with the logos of SGFI sponsors and local sponsors. The approval of all branding material should be taken by SGFI office in advance. Railway Station/Bus Stand/Airport and city should also be branded with the branding material of NSG.\n", normalFont));

            Weather_rec.Add(new Chunk("(25)   ", boldFont));
            Weather_rec.Add(new Chunk("Field of Play:-\n", titleFont3));
            Weather_rec.Add(new Chunk("          Field of Play should be of International level and as per the standards already been set by respective International Federation.\n", normalFont));

            Weather_rec.Add(new Chunk("(26)   ", boldFont));
            Weather_rec.Add(new Chunk("Remuneration to ITO/NTO/SGFI Staff and others:-\n", titleFont3));
            Weather_rec.Add(new Chunk("          The remuneration & TA/DA should be paid to ITO/NTO/SGFI Staff and others as per SGFI Financial guidelines.\n", normalFont));

            Weather_rec.Add(new Chunk("           (I)", normalFont2));
            Weather_rec.Add(new Chunk("If a team/Individual is absent by any reason/lacking/ fault and does not arrive in time at the site of the competition, the deposit on the entry fee, paid at the time of registration, will not be refunded at any cost.\n", normalFont));

            Weather_rec.Add(new Chunk("           (II)", normalFont2));
            Weather_rec.Add(new Chunk("A minimum number of 8 entries are required for each event, less than 8 entries will not be awarded certificate\r\n of merit. \n", normalFont));


            doc.Add(Weather_rec);

            doc.Add(new Paragraph("\n"));
            doc.Add(new Paragraph("\n"));


            table.DeleteBodyRows();
            table = new PdfPTable(2) { WidthPercentage = 100 };
            table.SetWidths(new float[] { 50f, 50f });

            table.AddCell(new PdfPCell(new Phrase("Date: 19/09/2024", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                Border = Rectangle.NO_BORDER,
                Padding = 2,
                Rowspan = 3
            });

            table.AddCell(new PdfPCell(new Phrase("Name & Designation", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                Border = Rectangle.NO_BORDER,
                Padding = 2
            });

            table.AddCell(new PdfPCell(new Phrase("Signature with Seal of Competent Authority", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                Border = Rectangle.NO_BORDER,
                Padding = 2
            });

            table.AddCell(new PdfPCell(new Phrase("State/UT/Unit", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                Border = Rectangle.NO_BORDER,
                Padding = 2
            });

            doc.Add(table);





            doc.Add(new Paragraph("\n"));
        }

        public class StateRequestsHeader : PdfPageEventHelper
        {
            public override void OnEndPage(PdfWriter writer, Document document)
            {
                Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 19, new BaseColor(163, 67, 123));
                Font subHeaderFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);
                Font contactFont = FontFactory.GetFont(FontFactory.HELVETICA, 9.5f);

                string imagePath = System.Web.Hosting.HostingEnvironment.MapPath("~/Media/638770343813090191.jpg");
                Image logo = Image.GetInstance(imagePath);
                logo.ScaleAbsolute(60f, 60f);

                // 6 Columns Table
                PdfPTable headerTable = new PdfPTable(6);
                headerTable.TotalWidth = document.PageSize.Width - 40;
                headerTable.SetWidths(new float[] { 10f, 20f, 20f, 10f, 10f, 30f }); // Adjusted widths
                headerTable.LockedWidth = true;

                // ===== LOGO CELL (Span 2 columns) =====
                PdfPCell logoCell = new PdfPCell(logo)
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 2,
                    Rowspan = 3
                };
                headerTable.AddCell(logoCell);


                headerTable.AddCell(new PdfPCell(new Phrase("School Games Federation of India", headerFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 4,
                    Rowspan = 1
                });

                headerTable.AddCell(new PdfPCell(new Phrase("Campus Radha Badha Inter College", contactFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 4,
                    Rowspan = 1
                });
                //headerTable.AddCell(textCell2);

                headerTable.AddCell(new PdfPCell(new Phrase("Shahganj, Agra - 282010 (U.P.) India", contactFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 4,
                    Rowspan = 1
                });


                headerTable.AddCell(new PdfPCell(new Phrase("Recognised by- Ministry of Youth Affairs & Sports, Govt. of India", contactFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 3
                });

                headerTable.AddCell(new PdfPCell(new Phrase("Tel: +91 0562-2211107, Mob: +91 9837885006", contactFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 3
                });

                headerTable.AddCell(new PdfPCell(new Phrase("Affiliated With International School Sports Federation,", contactFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 3
                });

                headerTable.AddCell(new PdfPCell(new Phrase("E-mail: infosgfi@sgfibharat.com, president@sgfibharat.com", contactFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 3
                });


                headerTable.AddCell(new PdfPCell(new Phrase("Asian School Sports Federation, Asian School Football Federation", contactFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 3
                });


                headerTable.AddCell(new PdfPCell(new Phrase("Website: www.sgfibharat.com", contactFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    Colspan = 3
                });


                // Add a green separator line
                LineSeparator lineHeader = new LineSeparator(2f, 100f, new BaseColor(12, 132, 52), Element.ALIGN_CENTER, -1);
                ColumnText.ShowTextAligned(writer.DirectContent, Element.ALIGN_CENTER, new Phrase(new Chunk(lineHeader)), document.PageSize.Width / 2, document.PageSize.Height - 120, 0);

                // Render header table on the PDF
                headerTable.WriteSelectedRows(0, -1, 20, document.PageSize.Height - 5, writer.DirectContent);







                /*PdfPCell textCell = new PdfPCell(headerText)
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    PaddingTop = 5,
                    PaddingBottom = 5,
                    Colspan = 2
                };
                headerTable.AddCell(textCell);*/


                /*// ===== LOGO CELL (Span 2 columns) =====
                PdfPCell logoCell = new PdfPCell(logo)
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    PaddingLeft = 10,
                    PaddingTop = 5,
                    PaddingBottom = 5,
                    Colspan = 3,
                    Rowspan=3,
                };
                headerTable.AddCell(logoCell);*/



            }

        }

        public class StateRequestsFooter : PdfPageEventHelper
        {
            public override void OnEndPage(PdfWriter writer, Document document)
            {
                PdfPTable footerTable = new PdfPTable(7);
                footerTable.TotalWidth = document.PageSize.Width - 80;
                float[] columnWidths = { 1f, 1f, 1f, 1f, 1f, 1f, 0.5f };
                footerTable.SetWidths(columnWidths);

                string[] logos = {
                    "~/Media/image11.jpeg",
                    "~/Media/image2.jpeg",
                    "~/Media/image3.png",
                    "~/Media/image4.jpg",
                    "~/Media/image5.png",
                    "~/Media/15384229492942.png"
                };

                //string imagePath = System.Web.Hosting.HostingEnvironment.MapPath("~/Media/PlayerList/638764558156394463.png");

                string imagePath = "";
                foreach (string logoPath in logos)
                {
                    imagePath = System.Web.Hosting.HostingEnvironment.MapPath(logoPath);
                    Image logo = Image.GetInstance(imagePath);
                    logo.ScaleAbsolute(50f, 50f);
                    PdfPCell logoCell = new PdfPCell(logo)
                    {
                        Border = Rectangle.NO_BORDER,
                        HorizontalAlignment = Element.ALIGN_CENTER,
                        PaddingBottom = 10f
                    };
                    footerTable.AddCell(logoCell);
                }
                PdfPCell pageNumberCell = new PdfPCell(new Phrase($"Page - {writer.PageNumber}", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    PaddingBottom = 10f
                };
                footerTable.AddCell(pageNumberCell);

                footerTable.WriteSelectedRows(0, -1, 40, 60, writer.DirectContent);
                //footerTable.WriteSelectedRows(0, -1, 40, 50, writer.DirectContent);
            }
        }
        #endregion

        #region  Certificate
        public ActionResult GenerateCertificate()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                Document document = new Document(PageSize.A4, 0, 0, 0, 0);
                PdfWriter writer = PdfWriter.GetInstance(document, stream);
                document.Open();

                string imagePath = System.Web.Hosting.HostingEnvironment.MapPath("~/Media/638770343813090191.jpg");
                Image logo = Image.GetInstance(imagePath);
                logo.ScaleAbsolute(60f, 60f);

                Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 28, new BaseColor(163, 67, 123));
                Font subHeaderFont = FontFactory.GetFont(FontFactory.HELVETICA, 15);
                Font contactFont = FontFactory.GetFont(FontFactory.HELVETICA, 9.5f);
                Font contentBLACK = FontFactory.GetFont(FontFactory.TIMES_ROMAN, 12, BaseColor.BLACK);
                Font contentRED = FontFactory.GetFont(FontFactory.TIMES_BOLDITALIC, 12, BaseColor.RED);

                PdfPTable headerTable = new PdfPTable(2);
                headerTable.TotalWidth = document.PageSize.Width - 40;
                headerTable.SetWidths(new float[] { 15f, 85f }); 
                headerTable.LockedWidth = true;

                PdfPCell logoCell = new PdfPCell(logo){Border = Rectangle.NO_BORDER,HorizontalAlignment = Element.ALIGN_LEFT,VerticalAlignment = Element.ALIGN_MIDDLE,PaddingTop = 10f};
                headerTable.AddCell(logoCell);


                headerTable.AddCell(new PdfPCell(new Phrase("School Games Federation of India", headerFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    VerticalAlignment = Element.ALIGN_MIDDLE,
                    
                });
                document.Add(headerTable);
                document.Add(new Paragraph("\n\n"));

                Paragraph title = new Paragraph($"Recognised by-Ministry of Youth Affairs & Sports,Govt Of India\n" +
                    $"Member:International School Sports Federation, Asian School Sports Federation, Asian School Football Federation", FontFactory.GetFont(FontFactory.HELVETICA, 9.5f));
                title.Alignment = Element.ALIGN_CENTER;
                document.Add(title);

                document.Add(new Paragraph("\n"));


                PdfPTable games = new PdfPTable(2);
                games.TotalWidth = document.PageSize.Width;
                games.SetWidths(new float[] { 75f, 25f });

                Paragraph PP = new Paragraph();
                PP.Add(new Chunk("68TH NATIONAL SCHOOL GAMES 2024-25\r\n", FontFactory.GetFont(FontFactory.HELVETICA, 15, BaseColor.DARK_GRAY)));
                PP.Add(new Chunk("TABLE TENNIS BOYS U-19\r\n", FontFactory.GetFont(FontFactory.HELVETICA, 15, BaseColor.RED)));
                PP.Add(new Chunk("LEH,LADAKH\r\n", FontFactory.GetFont(FontFactory.HELVETICA, 15, BaseColor.RED)));
                PP.Add(new Chunk("05-10-2024 to 08-10-2024\r\n", FontFactory.GetFont(FontFactory.HELVETICA, 15, BaseColor.RED)));
                PP.Alignment = Element.ALIGN_CENTER;

                games.AddCell(new PdfPCell(PP)
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    VerticalAlignment = Element.ALIGN_MIDDLE,

                });
                PdfPCell photogroph = new PdfPCell(logo) { Border = Rectangle.NO_BORDER, HorizontalAlignment = Element.ALIGN_CENTER, VerticalAlignment = Element.ALIGN_MIDDLE, PaddingTop = 10f };
                games.AddCell(photogroph);
                document.Add(games);

                document.Add(new Paragraph("\n\n"));

                headerTable.DeleteBodyRows();
                headerTable = new PdfPTable(4) { WidthPercentage = 80};
                headerTable.SetWidths(new float[] { 18f, 25f, 22f, 35f });

                headerTable.AddCell(CreateHeaderCell1("Type"));
                headerTable.AddCell(CreateHeaderCell1("Registration No."));
                headerTable.AddCell(CreateHeaderCell1("Certificate No."));
                headerTable.AddCell(CreateHeaderCell1("Organised By"));
                for (int i = 1; i <= 1; i++)
                {
                    headerTable.AddCell(CreateBodyCell("PRIORITY", Element.ALIGN_CENTER));
                    headerTable.AddCell(CreateBodyCell("20242529374733", Element.ALIGN_CENTER));
                    headerTable.AddCell(CreateBodyCell("293-4217-74733", Element.ALIGN_CENTER));
                    headerTable.AddCell(CreateBodyCell("Directorate of Youth Service & Sports, UT of Ladakh", Element.ALIGN_CENTER));
                }
                document.Add(headerTable);

                document.Add(new Paragraph("\n\n"));

                Paragraph Merit = new Paragraph($"Certificate of Merit", FontFactory.GetFont(FontFactory.HELVETICA, 25f,BaseColor.BLACK));
                Merit.Alignment = Element.ALIGN_CENTER;
                document.Add(Merit);

                document.Add(new Paragraph("\n\n"));



                Phrase phrase = new Phrase();
                Paragraph playerDetails = new Paragraph
                {
                    Alignment = Element.ALIGN_LEFT,
                    PaddingTop = 10f,

                };
                //var phrase = new Phrase(new Chunk("Awarded that Miss/Mr.: ", contentBLACK)).Append(new Chunk("KUSHAL CHOPDA", contentRED));



                phrase.Add(new Chunk($"Awarded that Miss/Mr.:", contentBLACK));
                phrase.Add(new Chunk($"    KUSHAL CHOPDA\n", contentRED));
                phrase.Add(new Chunk($"Father's name:", contentBLACK));
                phrase.Add(new Chunk($" Mr. PIYUSH CHOPDA\n", contentRED));
                phrase.Add(new Chunk($"DOB:    ", contentBLACK));
                phrase.Add(new Chunk($"22-04-2007", contentRED));
                phrase.Add(new Chunk($"      CLASS:", contentBLACK));
                phrase.Add(new Chunk($" XII\n", contentRED));
                phrase.Add(new Chunk($"Participated in the:    ", contentBLACK));
                phrase.Add(new Chunk($"68TH NATIONAL SCHOOL GAMES TABLE TENNIS CHAMPIONSHIP / TOURNAMENT 2024-25\n", contentRED));
                phrase.Add(new Chunk($"From State/UT/Unit: ", contentBLACK));
                phrase.Add(new Chunk($"MAHARASHTRA\n", contentRED));
                phrase.Add(new Chunk($"Has been declared position:", contentBLACK));
                phrase.Add(new Chunk($"    FIRST", contentRED));
                phrase.Add(new Chunk($"      In Event/Discipline:", contentBLACK));
                phrase.Add(new Chunk($" TABLE TENNIS\n", contentRED));
                phrase.Add(new Chunk($"And Achieved ", contentBLACK));
                phrase.Add(new Chunk($"   (TEAM)\n", contentRED));

                playerDetails.Add(phrase);
                playerDetails.IndentationRight = 60;
                playerDetails.IndentationLeft = 60;

                document.Add(playerDetails);

                document.Add(new Paragraph("\n\n\n\n"));

                headerTable.DeleteBodyRows();
                headerTable = new PdfPTable(2) { WidthPercentage = 80 };
                headerTable.SetWidths(new float[] { 50f, 50f });
                PdfPCell photogroph1 = new PdfPCell(logo) { Border = Rectangle.NO_BORDER, HorizontalAlignment = Element.ALIGN_RIGHT, VerticalAlignment = Element.ALIGN_MIDDLE,Rowspan=2 };

                headerTable.AddCell(new PdfPCell(new Phrase($"Date- {DateTime.Today.ToString("dd/MM/yyyy")}", contentBLACK)){ Border = Rectangle.NO_BORDER, BackgroundColor = BaseColor.WHITE,HorizontalAlignment = Element.ALIGN_LEFT});
                headerTable.AddCell(photogroph1);
                headerTable.AddCell(new PdfPCell(new Phrase($"Place- LEH", contentBLACK)){ Border = Rectangle.NO_BORDER, BackgroundColor = BaseColor.WHITE,HorizontalAlignment = Element.ALIGN_LEFT });
                
                document.Add(headerTable);




                PdfPTable footerTable = new PdfPTable(3);
                footerTable.TotalWidth = document.PageSize.Width - 80;
                float[] columnWidths = { 33.3f, 33.4f, 33.3f };
                footerTable.SetWidths(columnWidths);
                PdfContentByte cb = writer.DirectContent;

                cb.SetLineWidth(0.5f);
                cb.MoveTo(40, 100); // Starting point for left cell's line
                cb.LineTo(40 + (footerTable.TotalWidth / 4), 100); // Ending point for left cell's line
                cb.Stroke();
                // Left cell
                PdfPCell leftCell = new PdfPCell(new Phrase("Deepak Kumar (IAS)\nPresident", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    PaddingBottom = 10f
                };
                footerTable.AddCell(leftCell);

                cb.MoveTo(40 + (footerTable.TotalWidth / 4), 100); // Starting point for middle cell's line
                cb.LineTo(40 + 2 * (footerTable.TotalWidth / 4), 100); // Ending point for middle cell's line
                cb.Stroke();
                // Middle cell
                PdfPCell middleCell = new PdfPCell(new Phrase("Observer/Authority\nSGFI", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    PaddingBottom = 10f
                };
                footerTable.AddCell(middleCell);


                cb.MoveTo(40 + 2 * (footerTable.TotalWidth / 4), 100);
                cb.LineTo(40 + footerTable.TotalWidth, 100); 
                cb.Stroke();

                PdfPCell rightCell = new PdfPCell(new Phrase("Amarjeet K. Sharma\nWorking Secretary General", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER,
                    PaddingBottom = 10f
                };
                footerTable.AddCell(rightCell);

                footerTable.WriteSelectedRows(0, -1, 60, 60, writer.DirectContent);






                document.Close();

                byte[] bytes = stream.ToArray();
                return File(bytes, "application/pdf", "Certificate.pdf");
            }
        }

        #endregion
        public static PdfPTable MakePlayer(Players model)
        {
            Document doc = new Document(PageSize.A4);
            MemoryStream memStream = new MemoryStream();
            Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
            Font normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);

            PdfPTable HeaderTable1 = new PdfPTable(1)
            {
                WidthPercentage = 100
            };

            try
            {
                PdfWriter writer = PdfWriter.GetInstance(doc, memStream);
                doc.Open();

                Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 20, new BaseColor(163, 67, 123)); // #a3437b color
                Paragraph title = new Paragraph("SCHOOL GAMES FEDERATION OF INDIA", titleFont)
                {
                    Alignment = Element.ALIGN_CENTER
                };
                doc.Add(title);
                doc.Add(new Paragraph("\n"));

                HeaderTable1.SetWidths(new float[] { 30f, 40f, 30f });

                PdfPCell leftCell = new PdfPCell(new Phrase("ARCHERY U-17 Boys", normalFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_LEFT
                };

                PdfPCell centerCell = new PdfPCell(new Phrase("68TH NATIONAL SCHOOL GAMES (2024-25)", boldFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };

                PdfPCell rightCell = new PdfPCell(new Phrase("State: MANIPUR", normalFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT
                };

                HeaderTable1.AddCell(leftCell);
                HeaderTable1.AddCell(centerCell);
                HeaderTable1.AddCell(rightCell);

                doc.Add(HeaderTable1);
                //doc.Add(new Paragraph("\n")); // Add some spacing

                PdfPTable eventTable = new PdfPTable(1)
                {
                    WidthPercentage = 100
                };

                StringBuilder eventDetails = new StringBuilder();
                eventDetails.AppendLine("11.11.2024 To 12.11.2024, Nadiad\n");
                eventDetails.AppendLine("Organized By: Sports Authority of Gujarat, Gandhinagar\n");
                eventDetails.AppendLine("Under the aegis of School Games Federation of India");

                PdfPCell eventCell = new PdfPCell(new Phrase(eventDetails.ToString(), normalFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_CENTER
                };
                eventTable.AddCell(eventCell);
                doc.Add(eventTable);

                PdfPTable headerTable = new PdfPTable(2) { WidthPercentage = 100 };
                headerTable.SetWidths(new float[] { 75f, 25f });
                PdfPCell formHeaderCell = new PdfPCell(new Phrase("68TH NATIONAL SCHOOL GAMES (2024-25)", boldFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT
                };

                PdfPCell dateCell = new PdfPCell(new Phrase("Date: " + DateTime.Now.ToString("dd/MMM/yyyy"), normalFont))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = Element.ALIGN_RIGHT
                };
                headerTable.AddCell(formHeaderCell);
                headerTable.AddCell(dateCell);
                doc.Add(headerTable);





                PdfPTable table = new PdfPTable(9)
                {
                    WidthPercentage = 100
                };
                table.SetWidths(new float[] { 5f, 15f, 15f, 15f, 10f, 10f, 15f, 10f, 10f });

                string[] headers = { "SN", "Reg. No", "Name", "Father's Name", "DOB", "Cls", "Category", "School Name", "Photo" };

                foreach (string header in headers)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(header, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, BaseColor.WHITE)))
                    {
                        BackgroundColor = new BaseColor(0, 102, 204),
                        HorizontalAlignment = Element.ALIGN_CENTER
                    };
                    table.AddCell(cell);
                }


                int i = 1;
                foreach (DataRow player in model.dt.Rows)
                {
                    table.AddCell(new PdfPCell(new Phrase((++i).ToString(), normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                    table.AddCell(new PdfPCell(new Phrase(player["SGFIRegNo"].ToString(), normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                    table.AddCell(new PdfPCell(new Phrase(player["Name"].ToString(), normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                    table.AddCell(new PdfPCell(new Phrase(player["FatherName"].ToString(), normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                    table.AddCell(new PdfPCell(new Phrase(player["DateOfBirth"].ToString(), normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                    table.AddCell(new PdfPCell(new Phrase(player["PlayerClass"].ToString(), normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                    table.AddCell(new PdfPCell(new Phrase(player["GameName"].ToString(), normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                    table.AddCell(new PdfPCell(new Phrase(player["SchoolName"].ToString(), normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });

                    try
                    {
                        string imagePath = System.Web.Hosting.HostingEnvironment.MapPath("~/" + player["PlayerPhotograph"]);
                        if (System.IO.File.Exists(imagePath))
                        {
                            Image img = Image.GetInstance(imagePath);
                            img.ScaleAbsolute(40, 40);
                            PdfPCell imgCell = new PdfPCell(img) { HorizontalAlignment = Element.ALIGN_CENTER };
                            table.AddCell(imgCell);
                        }
                        else
                        {
                            table.AddCell(new PdfPCell(new Phrase("No Photo", normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                        }
                    }
                    catch
                    {
                        table.AddCell(new PdfPCell(new Phrase("No Photo", normalFont)) { HorizontalAlignment = Element.ALIGN_CENTER });
                    }
                }

                doc.Add(table);
                doc.Close();

            }
            catch (Exception ex)
            {

            }


            return HeaderTable1;
        }


    }
}

