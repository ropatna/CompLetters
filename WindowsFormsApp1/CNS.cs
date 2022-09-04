using System;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class CNS : Form
    {
        static DateTime date = DateTime.Now;
        string date_str = date.ToString("dd/MM/yyyy"); //CURRENT SYSTEM DATE
        SqlConnection sqlcon = new SqlConnection(connectionString: "Data Source=CBSEPAT\\SQLEXPRESS;Initial Catalog=LETTERS;Integrated Security=True"); //CONNECTION STRING
        public CNS()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand("", sqlcon);
            string database = Microsoft.VisualBasic.Interaction.InputBox("ENTER NAME OF DATABASE FROM WHICH LETTER HAS TO BE GENERATED", "INPUT DATABASE NAME", "hecomp2022c");
            if (String.IsNullOrEmpty(textBox3.Text))
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "]", sqlcon);
            }
            else
            {
                cmd = new SqlCommand("select * FROM [LETTERS].[dbo].[" + database + "] where cns_schno='" + textBox3.Text + "'", sqlcon);
            }
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                Document pdoc = new Document(PageSize.A4, 20f, 20f, 10f, 50f);
                PdfWriter pwriter = PdfWriter.GetInstance(pdoc, new FileStream("D:\\ABHISHEK\\WORK2022C\\PDF LETTERS\\CNS\\" + dr["cns_schno"].ToString() + "_CNS.pdf", FileMode.Create));
                var header = iTextSharp.text.Image.GetInstance("C:\\Users\\Acer\\source\\repos\\ropatna\\CompLetters\\WindowsFormsApp1\\images\\header.png");
                var footer = iTextSharp.text.Image.GetInstance("C:\\Users\\Acer\\source\\repos\\ropatna\\CompLetters\\WindowsFormsApp1\\images\\FOOTER.png");
                var rosign = iTextSharp.text.Image.GetInstance("C:\\Users\\Acer\\source\\repos\\ropatna\\CompLetters\\WindowsFormsApp1\\images\\rosignpng.png");
                var header2 = iTextSharp.text.Image.GetInstance("C:\\Users\\Acer\\source\\repos\\ropatna\\CompLetters\\WindowsFormsApp1\\images\\acceptance_header.png");
                header2.ScaleToFit(900f, 60f);
                header2.Alignment = 1;
                header.ScaleToFit(900f, 60f);
                header.ScaleToFit(900f, 60f);
                header.Alignment = 1;
                footer.ScaleToFit(880f, 55f);
                footer.SetAbsolutePosition(15, 10);
                footer.Alignment = 1;
                rosign.ScaleToFit(100f, 40f);
                rosign.SetAbsolutePosition(470, 125);
                pdoc.Open();
                pdoc.Add(header); //Adding Header
                pdoc.Add(footer); //Adding Foter
                iTextSharp.text.Font arial = FontFactory.GetFont("Arial", 12);
                iTextSharp.text.Font bold = FontFactory.GetFont(FontFactory.TIMES_BOLD, 12);
                pdoc.AddTitle("CNS Acceptance Letter");
                //
                Paragraph p = new Paragraph("===============================================================================\n");
                pdoc.Add(p);
                Paragraph p2 = new Paragraph("CONFIDENTIAL", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p2);
                Paragraph p3 = new Paragraph("*******************") { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p3);
                Paragraph p4 = new Paragraph(str: "No.:CBSE/RO(PTN)/CONF./NODAL/" + dr["cns_schno"].ToString() + "/COMP-2022/                                            Date:" + date_str+"\n\n") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p4);
                Paragraph p5 = new Paragraph(str: dr["CNSNAME"].ToString()+ "\nPRINCIPAL(School Code: " + dr["cns_schno"].ToString() + ")\n"  + "" + dr["cnsadd1"].ToString() + "\n" + dr["cnsadd2"].ToString() + "\n" + dr["cnsadd3"].ToString() + "\n" + dr["cnsadd4"].ToString() + " " + dr["cnsadd5"].ToString() + " - " + dr["cnspin"].ToString() + "\n\nSUBJECT : Appointment cum Intimation/Acceptance of Chief Nodal Supervisor (CNS) For Spot Evaluation for AISSE/AISSCE - Compartment Examination - 2022.\n\nSir/Madam,\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p5);
                Paragraph p6 = new Paragraph(str: "      This is to inform you that the Competent Authority of the Board is pleased to appoint you as Chief Nodal Supervisor at your school / nodal center to ensure proper supervision and timely completion of Evaluation of Answer books with perfect  accuracy. \n\nThe Compartment  Examination of AISSE/AISSCE 2022 is scheduled to be commenced from 23-08-2022 to 29-08-2022. The Chief Nodal Supervisor(CNS) would be a Principal of the school and Vice Principal/PGT be  appointed  as  Head Examiner  under his/her supervision. The  Chief Nodal Supervisor would supervise the  work of  Head Examiners appointed under him/her.  The Head  Examiner would do the Evaluation Work in the School of Chief Nodal Supervisor.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p6);
                Paragraph p7 = new Paragraph(str: "\n     The evaluation will be done at your school and the timing of evaluation will be 8 to 10 hours minimum for HE/AHE (Evaluation) and AHE (Coord)/Evaluators in accordance with the schedule given by the Board. Since evaluation is a time-bound activity, it is, therefore, desirable that adequate time be devoted by each teacher to complete the Evaluation not only in time but should be in a perfect and professional manner under each Chief Nodal Supervisor, Head -Examiners and around  12 - 16 Examiners are also being appointed with each Head Examiner.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p7);
                Paragraph p8 = new Paragraph(str: "\n      For evaluation work, one separate room is to be provided for each Head Examiner. The tentative dates for Evaluation are between 25-08-2022 to 29-08-2022. The AHE (Coord.) who will do coordination work of the Answer Books (uploading of marks) and AHE (Evaluation) will be appointed by the Head Examiners as per provisions/guidelines amongst PGTs and TGTs for classes XII and X respectively in consultation with you. ") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p8);
                pdoc.NewPage();
                pdoc.Add(header);
                pdoc.Add(footer);
                Paragraph p30 = new Paragraph(str: "      The Evaluation/Coordination work will be done strictly as per instructions given in the guidelines for Spot Evaluation.",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p30);
                Paragraph p31 = new Paragraph(str: "\n      The   delivery   of   Answer -Books will be made approximately after 04-05 days from the commencement of the Examination of  the   concerned  subject and evaluation shall be got started immediately on  receipt of Answer Books by you at Nodal Centre. The appointment as well as the information in this regard will be kept secret.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p31);
                Paragraph p32 = new Paragraph(str: "\n      As a Chief Nodal Supervisor, you will  have to  ensure that  the Head Examiners and AHE / Examiners working under your supervision are Evaluating the Answer Scripts strictly in accordance with the  Marking Scheme and leaving  no  scope  for  any allegations etc. except the real merit of the Examinees.  It may also be ensured that the AHE (Evaluation) / AHE (Coord) will perform their duties with utmost care and sense of responsibility and will see to it that evaluation/coordination work done by them is absolute without any error and prejudice.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p32);
                Paragraph p33 = new Paragraph(str: "\n      Rates Of Remuneration/Conveyance Is Admissible As Per Spot Guidelines Compartment Examination 2022.",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p33);
                Paragraph p9 = new Paragraph(str: "\n      FURTHER, AS PER PAST EXPERIENCE, IT HAS BEEN OBSERVED THAT QUALITY OF EVALUATION DONE AT SOME OF THE SPOT EVALUATION CENTRES WAS NOT FOUND SATISFACTORY AND ALSO NOT UPTO THE  DESIRED LEVEL AS PER THE MARKING SCHEME.  A LARGE NUMBER OF MISTAKE CASES WERE DETECTED DURING THE COURSE OF SCRUTINY ON ACCOUNT OF EVALUATION OF PREVIOUS EXAMINATIONS, WHICH WERE NOT DONE PROPERLY  &  RAISED    QUESTIONS  ON  THE  CREDIBILITY  ON  WORKING  OF  HEAD EXAMINERS AND EXAMINERS WHO PARTICIPATED IN THE EVALUATION WORK.", bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p9);
                Paragraph p10 = new Paragraph(str: "\n      As you are aware that students look at this examination as a final evaluation of their academic performance. The Competent Authority of the board has taken it seriously owing to large no. Of mistakes during re- evaluation. Also, from exam 2012 students/examinees may take photocopy of their answer sheets under RTI Act 2005 as per orders of the Hon'ble Supreme Court Of India also can get their evaluated Answer Books for opted subjects. Therefore, it is requested that proper attention towards evaluation should be given and answer books of the subject be evaluated in perfect manner strictly in accordance with the marking scheme.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p10);
                Paragraph p11 = new Paragraph(str: "      Apart from above, the examiners evaluating the answer books of the medium other than they are teaching, may have some difficulty in understanding the answer which may lead to  wrong  evaluation  of marks. Therefore, for doing full justice to  the students: all head examiners should strictly check and ensure that no answer book is evaluated by examiners who are teaching in the medium other than the one used in the answer book i.e. all answer books should be  checked by the  subject examiners of the same medium.",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p11);
                pdoc.NewPage();
                pdoc.Add(header);
                pdoc.Add(footer);
                Paragraph p34 = new Paragraph(str: "      Further, Head -Examiner should check the all Answer Books whether any  Answer Books of Physically Challenged children (Spasctic, Blind, Physically  Handicapped and Dyslexic candidates) have been erroneously received along with the   Answer Books of other candidates.   If the Answer Books of Physically Challenged candidates are found mixed with the Answer Books of other candidates, these be immediately returned to the undersigned without being  evaluated  through  sealed insured Speed Post parcel and information must be acknowledged to undersigned through email.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p34);
                Paragraph p12 = new Paragraph(str: "\n      After evaluation of all  the  Answer  Books  pertaining  to  the  H.E.s concerned be serialized in ascending order century wise and should be packed / sealed in respective Answer Book bags\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p12);
                Paragraph p35 = new Paragraph(str: "      It may  also  be  noted  that  the  Evaluation  is  the  most  important part  of the whole Examination system and it determines the  future career of the  students.  Therefore, you are  requested  to  take every possible care/effort to ensure objective and  judicious  Evaluation  to  safeguard the interest of the students and also to avoid any future complications / allegations. \n\n",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p35);
                Paragraph p13 = new Paragraph(str: "      A copy of the guidelines for  Spot Evaluation Compartment Examination 2022  will be e-mailed to you  in due course for your reference and strict compliance.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p13);
                Paragraph p14 = new Paragraph(str: "\n      You are also requested  to expedite the acceptance of  Head  Examiners appointed under your supervisions.") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p14);
                Paragraph p25 = new Paragraph(str: "\n      If  there is any change  on  account  of transfer, ward  appearing  in the  subject  etc.: the  same  should be informed immediately with replacement as per experience/eligibility so that necessary changes be done accordingly in time.", bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p25);
                Paragraph p26 = new Paragraph(str: "\n      Your acceptance of the aforesaid assignment in the enclosed proforma  duly completed in all  respects should be reached to the undersigned within 03 days from the issuance of this letter through Email to ropatna.cbse@nic.in and abcell.ropatna@cbseshiksha.in.\n\n",bold) { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p26);
                Paragraph p15 = new Paragraph(str: "      For any assistance with regard to above assignment: Email to abcell.ropatna@cbseshiksha.in mentioning complete details of CNS, School Number, School Name etc. from your school @cbseshiksha.in email id.\n\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p15);
                Paragraph p16 = new Paragraph(str: "      Yours faithfully,\n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p16);
                pdoc.Add(rosign);
                Paragraph p17 = new Paragraph(str: "\n(JAGADISH BARMAN) \nREGIONAL OFFICER\nRO PATNA, CBSE") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p17);
                pdoc.NewPage();
                pdoc.Add(header2); //Adding Header
                pdoc.Add(footer); //Adding Foter
                Paragraph p18 = new Paragraph(str: "ACCEPTANCE PROFORMA FOR CHEIF NODAL SUPERVISOR (CNS)\nFOR AISSE/AISSCE COMPARTMENT EXAMINATION - 2022\n-----------------------------------------------------------------------------------------\n(To be sent by Email or through Messenger)\nIMMEDIATE & CONFIDENTIAL", FontFactory.GetFont(FontFactory.TIMES_BOLD, 12)) { Alignment = Element.ALIGN_CENTER };
                pdoc.Add(p18);
                Paragraph p19 = new Paragraph(str: "Sch. No.:"+ dr["cns_schno"].ToString() + "\nDated:"+ date_str) { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p19);
                Paragraph p20 = new Paragraph(str: "THE REGIONAL OFFICER\nCENTRAL BOARD OF SECONDARY EDUCATION,\nREGIONAL OFFICE\nAMBIKA COMPLEX, BEHIND SBI COLONY NEAR BRAHMSTHAN\nSHEIKHPURA, BAILEY ROAD PATNA, BIHAR - 800014\n\nSir,\n\nWith  reference  to  your  Office  Letter  No.:CBSE/RO(PTN)/CONF./NODAL / " + dr["cns_schno"].ToString() + " /COMP 2022/ Dated  " + date_str + ". I  hereby  accept  to   act as Cheif Nodal Supervisor(CNS) for AISSE/AISSCE Compartment Examination - 2022.\n\n") { Alignment = Element.ALIGN_LEFT };
                pdoc.Add(p20);
                Paragraph p21 = new Paragraph(str: "      My appointment and any information that may come to my notice during the discharge of my duties as Chief Nodal Supervisor will be kept confidential.  I undertake to do this work with perfect efficiency and according to the instructions issued by the Board from time to time.\n\n      I CERTIFY THAT I HAVE NO NEAR  RELATION / WARD  INTENDING / APPEARING  IN THE SUBJECTS ALLOTTED AT MY NODAL CENTRE AT  THE  AFORESAID  EXAMINATION.  I  ALSO CERTIFY THAT I HAVE NOT  WRITTEN  ANY HELP  BOOK OR NOTES FOR THE  EXAMINATION OF THE BOARD.  I UNDERTAKE TO  COMPLETE  THE  WORK ENTRUSTED  TO  ME WITHIN THE STIPULATED TIME GIVEN IN THE LETTER OF APPOINTMENT.\n") { Alignment = Element.ALIGN_JUSTIFIED };
                pdoc.Add(p21);
                Paragraph p23 = new Paragraph(str: "Yours faithfully,     \n\nSignature ..................................\n\n") { Alignment = Element.ALIGN_RIGHT };
                pdoc.Add(p23);
                PdfPTable tbl3 = new PdfPTable(2);
                tbl3.WidthPercentage = 100f;
                tbl3.HorizontalAlignment = 1;
                tbl3.DefaultCell.Border = 0;
                tbl3.AddCell(new Phrase("Telephone (Office): " + dr["CNSSTD"].ToString() + "  " + dr["CNSteleres"].ToString() + "\n\n(Resi.if any):\n(Mobile No.): " + dr["CNSmobile"].ToString() + "\n\n\n\nMobile No. for Whatsapp:_______________________________"));
                tbl3.AddCell(new Phrase("Name: " + dr["CNSname"].ToString() + "\n\nSchool: " + dr["cnsadd1"].ToString() + "\n             " + dr["cnsadd2"].ToString() + "\n             " + dr["cnsadd3"].ToString() + "\n             " + dr["cnsadd4"].ToString() + "\n             (" + dr["cnsadd5"].ToString() + " - " + dr["cnspin"].ToString() + ")"));
                pdoc.Add(tbl3);
                //
                pdoc.Close();
            }
            MessageBox.Show("Voilla! Files Created.");
            sqlcon.Close();
        }
    }
}
