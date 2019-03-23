using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Admin
{
    public partial class ImportExcel : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnupload_Click(object sender, EventArgs e)
        {
            if (FileUploadProduct.HasFile)
            {
                try
                {
                    string path = string.Concat(Server.MapPath("~/Excel/" + FileUploadProduct.FileName));
                    FileUploadProduct.SaveAs(path);
                    string excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 8.0", path);
                    using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
                    {
                        connection.Open();
                        DataTable dtAOL = new DataTable();
                        DataTable dtCourse = new DataTable();
                        //connection.ConnectionString = excelConnectionString;
                        //AOL Master
                        OleDbCommand commandAOL = new OleDbCommand("select * from [AOL Master$]", connection);
                        DbDataReader drAOL = commandAOL.ExecuteReader();
                        dtAOL.Load(drAOL);
                        drAOL.Close();
                        commandAOL.Dispose();
                        var AOLSheet = dtAOL.AsEnumerable().Select(dataRow => new Models.University_Model
                        {
                            CountryName = Convert.ToString(dataRow.Field<dynamic>("Country")),
                            StateName = Convert.ToString(dataRow.Field<dynamic>("State")),
                            CityName = Convert.ToString(dataRow.Field<dynamic>("City")),
                            UniversityName = Convert.ToString(dataRow.Field<dynamic>("AOL Name")),
                            UniversityCode = Convert.ToString(dataRow.Field<dynamic>("AOL ID")),
                            LandMark = Convert.ToString(dataRow.Field<dynamic>("Location remarks")),
                            Affiliated = Convert.ToString(dataRow.Field<dynamic>("Affiliated with University for offering degree")),
                            GovApprovals = Convert.ToString(dataRow.Field<dynamic>("Govt Approvals")),
                            USP = Convert.ToString(dataRow.Field<dynamic>("USP")),
                            RankPosition = Convert.ToString(dataRow.Field<dynamic>("Rank / Position")),
                            PlacementRecord = Convert.ToString(dataRow.Field<dynamic>("Placement & Industrial Record")),
                            RecogAwardsAchiev = Convert.ToString(dataRow.Field<dynamic>("Recognitions / Awards / Achievements")),
                            InternationalCollaboration = Convert.ToString(dataRow.Field<dynamic>("International collaboration")),
                            CampusHighLights = Convert.ToString(dataRow.Field<dynamic>("Campus highlights")),
                            EstablishedYear = Convert.ToString(dataRow.Field<dynamic>("Year of establishment")),
                            History = Convert.ToString(dataRow.Field<dynamic>("History")),
                            RecentNews = Convert.ToString(dataRow.Field<dynamic>("Recently in News")),
                            Latitude = Convert.ToString(dataRow.Field<dynamic>("Latitude")),
                            Longitude = Convert.ToString(dataRow.Field<dynamic>("Longitude")),
                            Logo = Convert.ToString(dataRow.Field<dynamic>("Logo")),
                            MostPopular = Convert.ToString(dataRow.Field<dynamic>("MostPopular")),
                            Address = Convert.ToString(dataRow.Field<dynamic>("Address")),
                            WebSite = Convert.ToString(dataRow.Field<dynamic>("Website")),
                            ContactNo = Convert.ToString(dataRow.Field<dynamic>("Contact Number")),
                            EmailID = Convert.ToString(dataRow.Field<dynamic>("Email ID")),
                        }).ToList();

                        db_management.dba Objdba = new db_management.dba();
                        db_management.db_University ObjU = new db_management.db_University();
                        foreach (Models.University_Model University in AOLSheet)
                        {
                            if (!string.IsNullOrEmpty(University.UniversityCode))
                            {
                                University.UniversityID = Objdba.GetUiversityID(University.UniversityCode);
                                if (University.UniversityID == 0 && University.CityName != null)
                                {
                                    University.CountryID = Objdba.GetCountryID(University.CountryName);
                                    University.StateID = Objdba.GetStateID(University.CountryID, University.StateName);
                                    University.CityID = Objdba.GetCityID(University.CountryID, University.StateID, University.CityName);
                                }
                                else
                                {
                                    var Res = Objdba.GetUniversityLocationID(University.UniversityID);
                                    University.CountryID = Res.CountryID;
                                    University.StateID = Res.StateID;
                                    University.CityID = Res.CityID;
                                }
                                ObjU.AddUpdateUniversity(University);
                            }
                            else
                            {
                                MsgAlert.Text = "Updated upto empty field found.";
                                break;
                            }
                        }

                        //Course Master
                        OleDbCommand commandCourse = new OleDbCommand("select * from [Course Master$]", connection);
                        DbDataReader drCourse = commandCourse.ExecuteReader();
                        dtCourse.Load(drCourse);
                        drCourse.Close();
                        commandCourse.Dispose();
                        var CourseSheet = dtCourse.AsEnumerable().Select(dataRow => new Models.Course_Model
                        {
                            CountryName = Convert.ToString(dataRow.Field<dynamic>("Country")),
                            StateName = Convert.ToString(dataRow.Field<dynamic>("State")),
                            CityName = Convert.ToString(dataRow.Field<dynamic>("City")),
                            UniversityName = Convert.ToString(dataRow.Field<dynamic>("AOL Name")),
                            UniversityCode = Convert.ToString(dataRow.Field<dynamic>("AOL ID")),
                            CourseUniversityCode = Convert.ToString(dataRow.Field<dynamic>("Course ID")),
                            StreamName = Convert.ToString(dataRow.Field<dynamic>("Stream")),
                            LevelName = Convert.ToString(dataRow.Field<dynamic>("Degree Type")),
                            DegreeName = Convert.ToString(dataRow.Field<dynamic>("Degree Name")),
                            CourseName = Convert.ToString(dataRow.Field<dynamic>("Course Name")),
                            AnnualFee = Convert.ToString(dataRow.Field<dynamic>("Annual Fees")),
                            TotalFee = Convert.ToString(dataRow.Field<dynamic>("Total Fees")),
                            EligibleCriteria = Convert.ToString(dataRow.Field<dynamic>("Eligibility Criteria")),
                            CourseUSP = Convert.ToString(dataRow.Field<dynamic>("AOL Course USP")),
                            InternshipRecord = Convert.ToString(dataRow.Field<dynamic>("Placement / Internship Record")),
                            Ranking = Convert.ToString(dataRow.Field<dynamic>("Ranking (NIRF, any other pvt ranking)")),
                            AdmissionProcedure = Convert.ToString(dataRow.Field<dynamic>("Admission procedure")),
                            ApplicableEntranceExams = Convert.ToString(dataRow.Field<dynamic>("Applicable entrance exams")),
                            OverallAdmissionCriteria = Convert.ToString(dataRow.Field<dynamic>("Overall admission criteria")),
                            SpecialRequirement = Convert.ToString(dataRow.Field<dynamic>("Admission special requirement")),
                            ApplicationFee = Convert.ToString(dataRow.Field<dynamic>("Application fees")),
                            ApplicationLink = Convert.ToString(dataRow.Field<dynamic>("Application Link")),
                            CourseStartingDate = Convert.ToString(dataRow.Field<dynamic>("Course start date")),
                            IndianStudentsIntake = Convert.ToString(dataRow.Field<dynamic>("Indian students intake")),
                            ForeignStudentsIntake = Convert.ToString(dataRow.Field<dynamic>("Foreign students intake")),
                            SeatsReserved = Convert.ToString(dataRow.Field<dynamic>("Seat Reservation")),
                            IndustrialCollaboration = Convert.ToString(dataRow.Field<dynamic>("Industrial Collaboration")),
                            CurriculumSpeciality = Convert.ToString(dataRow.Field<dynamic>("Curriculum Speciality")),
                            PlacementSalaryRange = Convert.ToString(dataRow.Field<dynamic>("Placement Salary Range")),
                            PlacementAvgSalary = Convert.ToString(dataRow.Field<dynamic>("Placement Avg Salary")),
                            PlacedPercentage = Convert.ToString(dataRow.Field<dynamic>("Placed Percentage")),
                            PlacementLeadingCompany = Convert.ToString(dataRow.Field<dynamic>("Placement Leading Company")),
                            PlacementSpeciality = Convert.ToString(dataRow.Field<dynamic>("Placement speciality")),
                            IntenshipDetails = Convert.ToString(dataRow.Field<dynamic>("Internship Details")),
                            IntenshipStipend = Convert.ToString(dataRow.Field<dynamic>("Internship Stipend")),
                            CourseDuration = Convert.ToString(dataRow.Field<dynamic>("Course Duration in  years")),
                            AcademicUnitType = Convert.ToString(dataRow.Field<dynamic>("Academic Time Unit (Trimester / Semester / Yearly)")),
                            CurriculamDetails = Convert.ToString(dataRow.Field<dynamic>("Curriculam Details (Can provide link or separate hard copy)")),
                            AcademicFee = Convert.ToString(dataRow.Field<dynamic>("Academic Fee")),
                            RegistrationFee = Convert.ToString(dataRow.Field<dynamic>("Registration / Counselling Fee")),
                            HostelFee = Convert.ToString(dataRow.Field<dynamic>("Hostel Fee")),
                            TransportFee = Convert.ToString(dataRow.Field<dynamic>("Transport Fee")),
                            OtherFee = Convert.ToString(dataRow.Field<dynamic>("Other Fee")),
                            TotalCourseFee = Convert.ToString(dataRow.Field<dynamic>("Total Course Fee")),
                            ScholarshipAvailable = Convert.ToString(dataRow.Field<dynamic>("Scholarship available")),
                            ScholarshipDetails = Convert.ToString(dataRow.Field<dynamic>("Scholarship details")),
                            YearOfCourseStarted = Convert.ToString(dataRow.Field<dynamic>("Year of course started")),

                        }).ToList();

                        db_management.db_Course ObjCour = new db_management.db_Course();
                        foreach (Models.Course_Model CourseUniversity in CourseSheet)
                        {
                            if (!string.IsNullOrEmpty(CourseUniversity.UniversityCode))
                            {
                                CourseUniversity.UniversityID = Objdba.GetUiversityID(CourseUniversity.UniversityCode);
                                if (CourseUniversity.UniversityID != 0 && CourseUniversity.CourseName != null)
                                {
                                    CourseUniversity.StreamID = Objdba.GetStreamID(CourseUniversity.StreamName);
                                    CourseUniversity.LevelID = Objdba.GetLevelID(CourseUniversity.StreamID, CourseUniversity.LevelName);
                                    CourseUniversity.DegreeID = Objdba.GetDegreeID(CourseUniversity.StreamID, CourseUniversity.LevelID, CourseUniversity.DegreeName);
                                    CourseUniversity.CourseID = Objdba.GetCourseID(CourseUniversity.StreamID, CourseUniversity.LevelID, CourseUniversity.DegreeID, CourseUniversity.CourseName);
                                    ObjCour.AddUpdateCourseUniversity(CourseUniversity);
                                }
                            }
                            else
                            {
                                MsgAlert.Text = "Updated upto empty field found.";
                                break;
                            }
                        }
                        connection.Close();
                    }
                    MsgAlert.Text = "Uploaded successfully";
                }
                catch (Exception ex)
                {
                    MsgAlert.Text = ex.Message;
                }
            }
        }
    }
}