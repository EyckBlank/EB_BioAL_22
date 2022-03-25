///////////////////////////////////////////////////////////////////////////////////////////////////
// EB_BioAL_WPF
// zur Biologischen Umrechnung von Dosisverteilungen bei ExternalBeam
//
// abgeleitet von EB_Bio3D
//
// Versuch, Progressbar mit WPF hineinzubekommen
// Hilfe von Jan Suchotzki
// und Matt Schmidt
//
// 06.08.2021    1.0.0.4
// ConsoleFenster für ProgressAnzeige
// (von Marian Krüger)
//
// 9.8.2021 1.0.0.5
// VaginalZylinder
// GTV = PTV - VaginalZylinder
//
///////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Linq;
using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Linq.Expressions;
using System.Threading;
using System.Threading.Tasks;
//using EB_BioAL_WPF;
using System.Windows.Threading;
using EB_BioAL_WPF;



// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.5")]
[assembly: AssemblyFileVersion("1.0.0.5")]
[assembly: AssemblyInformationalVersion("1.05")]

// TODO: Uncomment the following line if the script requires write access.
[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        

        public Script()
        {
        }

       
        //#################

        double BioDose;

        const string BODY_ID1 = "BODY";
        const string BODY_ID2 = "Body";

        const string HIRN_ID = "Hirn";
        const string HIRNSTAMM_ID = "Hirnstamm";
        const string HIPPOCAMPUS_R_ID = "Hippocampus" + " " + "re";
        const string HIPPOCAMPUS_L_ID = "Hippocampus" + " " + "li";

        const string COCLEA_R_ID = "Coclea" + " " + "re";
        const string COCLEA_L_ID = "Coclea" + " " + "li";

        const string CHIASMA_ID = "Chiasma";
        const string SEHNERV_R_ID = "Sehnerv" + " " + "re";
        const string SEHNERV_L_ID = "Sehnerv" + " " + "li";

        const string LINSE_R_ID = "Linse" + " " + "re";
        const string LINSE_L_ID = "Linse" + " " + "li";

        const string PAROTIS_R_ID = "Parotis" + " " + "re";
        const string PAROTIS_L_ID = "Parotis" + " " + "li";

        const string SUBMANDI_R_ID = "Submandi" + " " + "re";
        const string SUBMANDI_L_ID = "Submandi" + " " + "li";

        const string MANDIBULA_ID = "Mandibula";
        const string CONSTRICTOR_ID = "Constrictor";

        const string MYELON_ID = "Myelon";

        const string PLEXUS_R_ID = "Plexus" + " " + "re";
        const string PLEXUS_L_ID = "Plexus" + " " + "li";

        const string LUNGE_R_ID = "Lunge" + " " + "re";
        const string LUNGE_L_ID = "Lunge" + " " + "li";

        const string HERZ_ID = "Herz";
        const string RIVA_ID = "Riva";

        const string BRUST_R_ID = "Brust" + " " + "re";
        const string BRUST_L_ID = "Brust" + " " + "li";

        const string OESOPHAGUS_ID = "Oesophagus";
        const string LEBER_ID = "Leber";

        const string NIERE_R_ID = "Niere" + " " + "re";
        const string NIERE_L_ID = "Niere" + " " + "li";

        const string DARM_ID = "Darm";
        const string REKTUM_ID = "Rektum";
        const string BLASE_ID = "Blase";

        const string GTV_ID = "GTV";
        const string VZ_ID = "VaginalZylinder";


        const string SCRIPT_NAME = "Bio IsodosenPlan Script";

        class abListD
        {
            internal double Body;
            internal double Hirn;
            internal double Hirnstamm;
            internal double Hippocampus;
            internal double Coclea;
            internal double Chiasma;
            internal double Sehnerv;
            internal double Linse;
            internal double Parotis;
            internal double Submandi;
            internal double Mandibula;
            internal double Constrictor;
            internal double Myelon;
            internal double Plexus;
            internal double Lunge;
            internal double Herz;
            internal double Riva;
            internal double Brust;
            internal double Oesophagus;
            internal double Leber;
            internal double Niere;
            internal double Darm;
            internal double Rektum;
            internal double Blase;
            internal double Gtv;
        }
        class abListS
        {
            internal string Body;
            internal string Hirn;
            internal string Hirnstamm;
            internal string Hippocampus;
            internal string Coclea;
            internal string Chiasma;
            internal string Sehnerv;
            internal string Linse;
            internal string Parotis;
            internal string Submandi;
            internal string Mandibula;
            internal string Constrictor;
            internal string Myelon;
            internal string Plexus;
            internal string Lunge;
            internal string Herz;
            internal string Riva;
            internal string Brust;
            internal string Oesophagus;
            internal string Leber;
            internal string Niere;
            internal string Darm;
            internal string Rektum;
            internal string Blase;
            internal string Gtv;
        }


        //-----------------------------------------------------------------------------------------
        // public void Execute(ScriptContext context, Window window) if a window should be shown
        public void Execute(ScriptContext context)
            ////public void Execute(ScriptContext context, Window window)
        {

            string PatLName = "";
            string PatFName = "";
            string PatID = "";
            string sBody = "";

            string bPlanPTV = "";
            string bPlanPrescr = "";

            double GD;
            double ED;
            double N;   //GD=ED*N
            double PI;  //prescr.isodose

            string setupDateipfad = Directory.GetCurrentDirectory();
            string debugDateipfad;
            setupDateipfad = setupDateipfad + "/";
            debugDateipfad = setupDateipfad;
            //MessageBox.Show(setupDateipfad);

            List<string> gtv = new List<string>();
            List<Structure> gtvls = new List<Structure>();

            abListS abS = new abListS();
            abListD abD = new abListD();

           
            abS = SetupExcel("EB_BioAL", "OGD Neuruppin",
                       setupDateipfad, "EB_BioAL_ini.xlsx",
                       debugDateipfad, "EB_BioAL.xlsx",
                       "Department", "Folders", "Organs",
                       "ab_Values", "Techn");

          

            abD.Body        = Convert.ToDouble(abS.Body);

            abD.Body        = Convert.ToDouble(abS.Body);
            abD.Hirn        = Convert.ToDouble(abS.Hirn);
            abD.Hirnstamm   = Convert.ToDouble(abS.Hirnstamm);
            abD.Hippocampus = Convert.ToDouble(abS.Hippocampus);
            abD.Coclea      = Convert.ToDouble(abS.Coclea);
            abD.Chiasma     = Convert.ToDouble(abS.Chiasma);
            abD.Sehnerv     = Convert.ToDouble(abS.Sehnerv);
            abD.Linse       = Convert.ToDouble(abS.Linse);

            abD.Parotis     = Convert.ToDouble(abS.Parotis);
            abD.Submandi    = Convert.ToDouble(abS.Submandi);
            abD.Mandibula   = Convert.ToDouble(abS.Mandibula);
            abD.Constrictor = Convert.ToDouble(abS.Constrictor);
            abD.Myelon      = Convert.ToDouble(abS.Myelon);

            abD.Plexus      = Convert.ToDouble(abS.Plexus);
            abD.Lunge       = Convert.ToDouble(abS.Lunge);
            abD.Herz        = Convert.ToDouble(abS.Herz);
            abD.Riva        = Convert.ToDouble(abS.Riva);
            abD.Brust       = Convert.ToDouble(abS.Brust);

            abD.Oesophagus  = Convert.ToDouble(abS.Oesophagus);
            abD.Leber       = Convert.ToDouble(abS.Leber);
            abD.Niere       = Convert.ToDouble(abS.Niere);
            abD.Darm        = Convert.ToDouble(abS.Darm);
            abD.Rektum      = Convert.ToDouble(abS.Rektum);
            abD.Blase       = Convert.ToDouble(abS.Blase);
            abD.Gtv         = Convert.ToDouble(abS.Gtv);

            // patient and plan context
            Patient Pat = context.Patient;
            PatLName = context.Patient.LastName.ToString();
            PatFName = context.Patient.FirstName.ToString();
            PatID = context.Patient.Id.ToString();

            context.Patient.BeginModifications();

            Course eCourse = context.Course;
            BrachyPlanSetup bPlan = context.BrachyPlanSetup;   // sonst kann ich BrachyPlan nicht auslesen
                       
            if (context.Patient == null || context.StructureSet == null)
            {
                MessageBox.Show("Please load a patient, 3D image, and structure set before running this script.", SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
                return;
            }
            StructureSet ss = context.StructureSet;

            foreach (Structure s in ss.Structures)
            {
                if (s.Id.Length > 3)
                {
                    if (s.Id == BODY_ID1 | s.Id == BODY_ID2)  // because upper-/lowercase
                    {
                        sBody = s.Id;
                    }
                    if (s.Id.Substring(0, 3) == "GTV")
                    {
                        gtv.Add(s.Id);
                        gtvls.Add(s);
                    }
                }
            }

            ExternalPlanSetup plan = eCourse.AddExternalPlanSetup(ss);  // nur als leerer Container weil ich sonst EvaluationPlanSetup nicht erzeugen kann

            // ePlanName
            string bPlanId = bPlan.Id;
            int startIndex = 0;
            int length = 10;
            string zName = bPlan.Name;
            int zLength = bPlanId.Length;
            if (zLength > 10)
            {
                length = 10;
            }
            else
            {
                length = zLength;
            }

            // ePlan erzeugen
            ExternalPlanSetup ePlan = eCourse.AddExternalPlanSetupAsVerificationPlan(ss, plan);  // Umweg über ExternalPlanSetup plan  da ich den ePlan sonst nicht an den Course bekomme     var
            EvaluationDose eDose = ePlan.CreateEvaluationDose();
            ePlan.CopyEvaluationDose(bPlan.Dose);         // to fit dimensions of dose matrices  dose und eDose !

            string ePlanId = bPlanId.Substring(startIndex, length) + "-A2";
            ePlan.Id = ePlanId;

            //eCourse.RemovePlanSetup(plan);

            MessageBox.Show("VeriPlan angelegt", SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);

            if (bPlan == null)
                return;

            int bFrakt = bPlan.NumberOfFractions.Value;

            bPlanPTV = bPlan.TargetVolumeID;

            GD = bPlan.TotalDose.Dose;
            ED = bPlan.DosePerFraction.Dose;
            N = (double)bPlan.NumberOfFractions.Value;
            PI = bPlan.TreatmentPercentage * 100;

            //ePlan.TotalDose.Dose = GD;
            //ePlan.DosePerFraction.Dose = ED;
            //ePlan.NumberOfFractions.Value = N;
            //ePlan.TreatmentPercentage = PI/100;


            string sGD = GD.ToString("F2") + bPlan.TotalDose.Unit;
            string sED = ED.ToString("F2");
            string sN  = N.ToString("F0");
            string sPI = PI.ToString("F2");

            bPlanPrescr = sED + " x " + sN + " = " + sGD + " (" + sPI + "%)";

            //=========================
            // Find the  structures
            //=========================

            // find Body 
            Structure body = ss.Structures.FirstOrDefault(x => x.Id == sBody);  // wegen Gro�- oder Kleinschreibung

            // find Hirn (brain)
            Structure hirn = ss.Structures.FirstOrDefault(x => x.Id == HIRN_ID);

            // find Hirnstamm (brainstem)
            Structure hirnstamm = ss.Structures.FirstOrDefault(x => x.Id == HIRNSTAMM_ID);

            // find Hippocampus_re (hippocampus right)
            Structure hippocampus_re = ss.Structures.FirstOrDefault(x => x.Id == HIPPOCAMPUS_R_ID);

            // find Hippocampus_li (hippocampus left)
            Structure hippocampus_li = ss.Structures.FirstOrDefault(x => x.Id == HIPPOCAMPUS_L_ID);

            // find Coclea_re (coclea right)
            Structure coclea_re = ss.Structures.FirstOrDefault(x => x.Id == COCLEA_R_ID);

            // find Coclea_li (coclea left)
            Structure coclea_li = ss.Structures.FirstOrDefault(x => x.Id == COCLEA_L_ID);

            // find Chiasma (chiasma opticus)
            Structure chiasma = ss.Structures.FirstOrDefault(x => x.Id == CHIASMA_ID);

            // find Sehnerv_re (nervus opticus right)
            Structure sehnerv_re = ss.Structures.FirstOrDefault(x => x.Id == SEHNERV_R_ID);

            // find Sehnerv_li (nervus opticus lrft)
            Structure sehnerv_li = ss.Structures.FirstOrDefault(x => x.Id == SEHNERV_L_ID);

            // find Linse_re (lens right)
            Structure linse_re = ss.Structures.FirstOrDefault(x => x.Id == LINSE_R_ID);

            // find Linse_li (lens left)
            Structure linse_li = ss.Structures.FirstOrDefault(x => x.Id == LINSE_L_ID);

            // find Parotis_re (gland parotis right)
            Structure parotis_re = ss.Structures.FirstOrDefault(x => x.Id == PAROTIS_R_ID);

            // find Parotis_li (gland parotis left)
            Structure parotis_li = ss.Structures.FirstOrDefault(x => x.Id == PAROTIS_L_ID);

            // find Submandi_re (gland submand right)
            Structure submandi_re = ss.Structures.FirstOrDefault(x => x.Id == SUBMANDI_R_ID);

            // find Submandi_li (gland (submand left)
            Structure submandi_li = ss.Structures.FirstOrDefault(x => x.Id == SUBMANDI_L_ID);

            // find Mandibula (mandibula)
            Structure mandibula = ss.Structures.FirstOrDefault(x => x.Id == MANDIBULA_ID);

            // find Constrictoor (muscule constrictor pharyngialis))
            Structure constrictor = ss.Structures.FirstOrDefault(x => x.Id == CONSTRICTOR_ID);

            // find Myelon (myelon)
            Structure myelon = ss.Structures.FirstOrDefault(x => x.Id == MYELON_ID);

            // find Plexus_re (nervus plexus right)
            Structure plexus_re = ss.Structures.FirstOrDefault(x => x.Id == PLEXUS_R_ID);

            // find Plexus_li (nervus plexus left)
            Structure plexus_li = ss.Structures.FirstOrDefault(x => x.Id == PLEXUS_L_ID);

            // find Lunge_re (lung right)
            Structure lunge_re = ss.Structures.FirstOrDefault(x => x.Id == LUNGE_R_ID);

            // find Lunge_li (lung right)
            Structure lunge_li = ss.Structures.FirstOrDefault(x => x.Id == LUNGE_L_ID);

            // find Herz (heart)
            Structure herz = ss.Structures.FirstOrDefault(x => x.Id == HERZ_ID);

            // find Riva (v. riva)
            Structure riva = ss.Structures.FirstOrDefault(x => x.Id == RIVA_ID);

            // find Brust_re (breast right)
            Structure brust_re = ss.Structures.FirstOrDefault(x => x.Id == BRUST_R_ID);

            // find Brust_li (breast left)
            Structure brust_li = ss.Structures.FirstOrDefault(x => x.Id == BRUST_L_ID);

            // find Oesophagus (oesophagus)
            Structure oesophagus = ss.Structures.FirstOrDefault(x => x.Id == OESOPHAGUS_ID);

            // find Leber (liver)
            Structure leber = ss.Structures.FirstOrDefault(x => x.Id == LEBER_ID);

            // find Niere_re (kidney right)
            Structure niere_re = ss.Structures.FirstOrDefault(x => x.Id == NIERE_R_ID);

            // find Niere_li (kidney left)
            Structure niere_li = ss.Structures.FirstOrDefault(x => x.Id == NIERE_L_ID);

            // find Darm (bowel)
            Structure darm = ss.Structures.FirstOrDefault(x => x.Id == DARM_ID);

            // find Rektum (rectum)
            Structure rektum = ss.Structures.FirstOrDefault(x => x.Id == REKTUM_ID);

            // find Blase (bladder)
            Structure blase = ss.Structures.FirstOrDefault(x => x.Id == BLASE_ID);

            // find VaginalZylinder (vaginalzylinder)
            Structure vz = ss.Structures.FirstOrDefault(x => x.Id == VZ_ID);

            // find GTV
            // how can I find all GTVs of structureset ?
            // Structure gtv1 = ss.Structures.FirstOrDefault(x => x.Id.Substring(0, 3) == GTV_ID);


            //=======================================
            // calculate the transformation parameters
            //=======================================

            var bDose = bPlan.Dose;

            int eX = 0;  // for debugger
            int eY = 0;
            int eZ = 0;

            double erX = 0;  // for debugger
            double erY = 0;
            double erZ = 0;

            int dSizeIX = bDose.XSize;
            int dSizeIY = bDose.YSize;
            int dSizeIZ = bDose.ZSize;     // dieses in anderen Programmen prüfen  Y zu Z

            double doX = bDose.Origin.x;
            double doY = bDose.Origin.y;
            double doZ = bDose.Origin.z;

            double dsizeX = bDose.XSize;
            double dsizeY = bDose.YSize;
            double dsizeZ = bDose.ZSize;

            double dresX = bDose.XRes;
            double dresY = bDose.YRes;
            double dresZ = bDose.ZRes;

            int[,]    zbuffer = new int[bDose.XSize, bDose.YSize];
            double[,] dbuffer = new double[bDose.XSize, bDose.YSize];
            double[,] ebuffer = new double[bDose.XSize, bDose.YSize];
            double[]  dbu     = new double[bDose.ZSize];
            int[]     dbi     = new int[bDose.ZSize];

            double[] xm = new double[bDose.XSize];
            double[] ym = new double[bDose.YSize];
            double[] zm = new double[bDose.ZSize];

            double fraktDose = bPlan.DosePerFraction.Dose;
            int fractNumber = (int)bPlan.NumberOfFractions;
            DoseValue ee = new DoseValue(fraktDose, DoseValue.DoseUnit.Gy);  // prepare fraction dose for new ePlan 
            ePlan.SetPrescription(fractNumber, ee, 1.0);               // here set fraction number and fraction dose of ePlan in ARIA 

            // ConsoleAllocator (MK)
            ConsoleAllocator.ShowConsoleWindow();
            Console.SetWindowSize(50, 6);
            Console.WriteLine("    ");
            Console.Write("    Schicht 0   von   " + bDose.ZSize.ToString() + "   ");

            VVector VV = new VVector(0, 0, 0);
            if (bDose != null)
            {
                // Durchlauf der Schichten
                for (int zi = 0; zi < bDose.ZSize; zi++)
                {
                    zm[zi] = zi * dresZ + doZ;
                    bDose.GetVoxels(zi, zbuffer);   // here read a dose layer

                    // indicate Progress ConsoleAllocator (MK)
                    Console.Write("\r" + "    Schicht " + zi.ToString() + "   von   " + bDose.ZSize.ToString() + "   ");

                    // 2D Durchlauf eines Bildes
                    for (int yi = 0; yi < bDose.YSize; yi++)
                    {
                        ym[yi] = yi * dresY + doY;

                        for (int xi = 0; xi < bDose.XSize; xi++)
                        {
                            xm[xi] = xi * dresX + doX;  // calculate carthesian points of matrix points for PointInsideSegment
                            VV[0] = xm[xi];
                            VV[1] = ym[yi];
                            VV[2] = zm[zi];

                            // IsPointInsideSegment for all entities 
                            // (long list)
                            // EQD2 conversion by a/b values
                            // 
                            // imageMatrix und doseMatrix have different dimensions
                            // IsPointInsideSegment must have carthesian coordinates
                            // x, y, z must converted by Xres, Yres, Zres and Origin 

                            dbuffer[xi, yi] = bDose.VoxelToDoseValue(zbuffer[xi, yi]).Dose * GD / 100;
                            ebuffer[xi, yi] = 0;

                            // Body
                            if (body.IsPointInsideSegment(VV))
                            {
                                if (body != null)
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Body) / (2 + abD.Body) * dbuffer[xi, yi];
                                }
                            }

                            // Hirn (brain)
                            if (hirn != null)
                            {
                                if (hirn.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Hirn) / (2 + abD.Hirn) * dbuffer[xi, yi];
                                }
                            }

                            // Hirnstamm (brainstem)
                            if (hirnstamm != null)
                            {
                                if (hirnstamm.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Hirnstamm) / (2 + abD.Hirnstamm) * dbuffer[xi, yi];
                                }
                            }

                            // Hippocampus_re (hippocampus right)
                            if (hippocampus_re != null)
                            {
                                if (hippocampus_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Hippocampus) / (2 + abD.Hippocampus) * dbuffer[xi, yi];
                                }
                            }
                            // Hippocampus_li (hippocampus left)
                            if (hippocampus_li != null)
                            {
                                if (hippocampus_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Hippocampus) / (2 + abD.Hippocampus) * dbuffer[xi, yi];
                                }
                            }

                            // Coclea_re (coclea right)
                            if (coclea_re != null)
                            {
                                if (coclea_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Coclea) / (2 + abD.Coclea) * dbuffer[xi, yi];
                                }
                            }
                            // Coclea_li (coclea left)
                            if (coclea_li != null)
                            {
                                if (coclea_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Coclea) / (2 + abD.Coclea) * dbuffer[xi, yi];
                                }
                            }

                            // Chiasma (chiasma)
                            if (chiasma != null)
                            {
                                if (chiasma.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Chiasma) / (2 + abD.Chiasma) * dbuffer[xi, yi];
                                }
                            }

                            // Sehnerv_re (nervus opticus right)
                            if (sehnerv_re != null)
                            {
                                if (sehnerv_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Sehnerv) / (2 + abD.Sehnerv) * dbuffer[xi, yi];
                                }
                            }
                            // Sehnerv_li (nervus opticus left)
                            if(sehnerv_li != null)
                            {
                                if (sehnerv_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Sehnerv) / (2 + abD.Sehnerv) * dbuffer[xi, yi];
                                }
                            }

                            // Linse_re (lens right)
                            if (linse_re != null)
                            {
                                if (linse_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Linse) / (2 + abD.Linse) * dbuffer[xi, yi];
                                }
                            }
                            // Linse_li (lens left)
                            if (linse_li != null)
                            {
                                if (linse_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Linse) / (2 + abD.Linse) * dbuffer[xi, yi];
                                }
                            }

                            // Parotis_re (gland parotis right)
                            if (parotis_re != null)
                            {
                                if (parotis_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Parotis) / (2 + abD.Parotis) * dbuffer[xi, yi];
                                }
                            }
                            // Parotis_li (gland parotis left)
                            if (parotis_li != null)
                            {
                                if (parotis_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Parotis) / (2 + abD.Parotis) * dbuffer[xi, yi];
                                }
                            }

                            // Submandi_re (gland submand right)
                            if (submandi_re != null)
                            {
                                if (submandi_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Submandi) / (2 + abD.Submandi) * dbuffer[xi, yi];
                                }
                            }
                            // Submandi_li (gland submand left)
                            if (submandi_li != null)
                            {
                                if (submandi_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Submandi) / (2 + abD.Submandi) * dbuffer[xi, yi];
                                }
                            }

                            // Mandibula (mandibule)
                            if (mandibula != null)
                            {
                                if (mandibula.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Mandibula) / (2 + abD.Mandibula) * dbuffer[xi, yi];
                                }
                            }

                            // Constrictor (muscule constrictor pharyngialis)
                            if (constrictor != null)
                            {
                                if (constrictor.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Constrictor) / (2 + abD.Constrictor) * dbuffer[xi, yi];
                                }
                            }

                            // Myelon (myelon)
                            if (myelon != null)
                            {
                                if (myelon.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Myelon) / (2 + abD.Myelon) * dbuffer[xi, yi];
                                }
                            }

                            // Plexus_re (nervus plexus right)
                            if (plexus_re != null)
                            {
                                if (plexus_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Plexus) / (2 + abD.Plexus) * dbuffer[xi, yi];
                                }
                            }
                            // Plexus_li (nervus plexus left)
                            if (plexus_li != null)
                            {
                                if (plexus_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Plexus) / (2 + abD.Plexus) * dbuffer[xi, yi];
                                }
                            }

                            // Lunge_re (lung right)
                            if (lunge_re != null)
                            {
                                if (lunge_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Lunge) / (2 + abD.Lunge) * dbuffer[xi, yi];
                                }
                            }
                            // Lunge_li (lung left)
                            if (lunge_li != null)
                            {
                                if (lunge_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Lunge) / (2 + abD.Lunge) * dbuffer[xi, yi];
                                }
                            }

                            // Herz (heart)
                            if (herz != null)
                            {
                                if (herz.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Herz) / (2 + abD.Herz) * dbuffer[xi, yi];
                                }
                            }

                            // Riva (v. riva)
                            if (riva != null)
                            {
                                if (riva.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Riva) / (2 + abD.Riva) * dbuffer[xi, yi];
                                }
                            }

                            // Brust_re (breast right)
                            if (brust_re != null)
                            {
                                if (brust_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Brust) / (2 + abD.Brust) * dbuffer[xi, yi];
                                }
                            }
                            // Brust_li (breast left)
                            if (brust_li != null)
                            {
                                if (brust_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Brust) / (2 + abD.Brust) * dbuffer[xi, yi];
                                }
                            }

                            // Oesophagus (oesophagus)
                            if (oesophagus != null)
                            {
                                if (oesophagus.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Oesophagus) / (2 + abD.Oesophagus) * dbuffer[xi, yi];
                                }
                            }

                            // Leber (liver)
                            if (leber != null)
                            {
                                if (leber.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Leber) / (2 + abD.Leber) * dbuffer[xi, yi];
                                }
                            }

                            // Niere_re (kidney right)
                            if (niere_re != null)
                            {
                                if (niere_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Niere) / (2 + abD.Niere) * dbuffer[xi, yi];
                                }
                            }
                            // Niere_li (kidney left)
                            if (niere_li != null)
                            {
                                if (niere_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Niere) / (2 + abD.Niere) * dbuffer[xi, yi];
                                }
                            }

                            // Darm (bowel / gut)
                            if (darm != null)
                            {
                                if (darm.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Darm) / (2 + abD.Darm) * dbuffer[xi, yi];
                                }
                            }

                            // Rektum (rectum)
                            if (rektum != null)
                            {
                                if (rektum.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Rektum) / (2 + abD.Rektum) * dbuffer[xi, yi];
                                }
                            }

                            // Blase (bladder)
                            if (blase != null)
                            {
                                if (blase.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Blase) / (2 + abD.Blase) * dbuffer[xi, yi];
                                }
                            }

                            // GTV list (gtvls)
                            if (gtvls != null)
                            {
                                foreach (Structure s1 in gtvls)
                                {
                                    if (s1.IsPointInsideSegment(VV))
                                    {
                                        ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Gtv) / (2 + abD.Gtv) * dbuffer[xi, yi];
                                    }
                                }
                            }

                            // VaginalZylinder (vaginalzylinder)
                            if (vz != null)
                            {
                                if (vz.IsPointInsideSegment(VV))
                                {
                                    // ebuffer[xi, yi] = (dbuffer[xi, yi] / bFrakt + abD.Blase) / (2 + abD.Blase) * dbuffer[xi, yi];
                                    ebuffer[xi, yi] = 0;
                                }
                            }

                            // all ebuffer reconvert into relative values
                            ebuffer[xi, yi] = ebuffer[xi, yi] * 100 / GD;
                            if (ebuffer[xi, yi] < 0)
                            {
                                ebuffer[xi, yi] = 0;
                            }
                            //ebuffer[xi, yi] in Applikatornähe;  
                            if (ebuffer[xi, yi] > 6000)   // 1500Gy
                            {
                                ebuffer[xi, yi] = 6000;
                            }
                            // zum Test ob er richtig rechnet oder die Indizes überlaufen
                            DoseValue eee = new DoseValue(ebuffer[xi, yi], DoseValue.DoseUnit.Percent);

                            zbuffer[xi, yi] = (int)eDose.DoseValueToVoxel(eee);

                        }
                    }

                   
                    // dbu[zi] = ebuffer[90, 54];    // for Excel debugging
                                                     //dbi[zi] = zbuffer[90, 54];

                    // for monitoring the matrix dimensions
                    eX = eDose.XSize;
                    eY = eDose.YSize;
                    eZ = eDose.ZSize;
                    erX = eDose.XRes;
                    erY = eDose.YRes;
                    erZ = eDose.ZRes;

                    eDose.SetVoxels(zi, zbuffer);

                }

                ConsoleAllocator.HideConsoleWindow();
                
                //----------------------------------------------------------------------------------
                // call of Debugging Excel
                //----------------------------------------------------------------------------------

            }
           

        }


        //################################################################################################################  
        //----------------------------------------------------------------------------------
        // Debugging Excel
        //----------------------------------------------------------------------------------
        private void UpdateExcel(string S1, string S2, string S3,
                                string S4, string S5, string S6,
                                string S11, string S12, string S13,
                                string S31, string S32, string S33)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open("Q:/ESAPI/Plugins/EB_BioAL/EB_BioAL.xlsx");
                oSheet = String.IsNullOrEmpty("Tab1") ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["Tab1"];

                //.................................................................................

                // read the row counter in Excel
                string sReihe = oSheet.Cells[1, 1].Value == null ? "-" : oSheet.Cells[1, 1].Value.ToString();
                int iReihe = Convert.ToInt32(sReihe);
                iReihe = iReihe + 1;

                // dresX
                //oSheet.Cells[iReihe, 1] = S1;
                //oSheet.Cells[iReihe, 2] = S2;
                //oSheet.Cells[iReihe, 3] = S3;
                // doX
                //oSheet.Cells[iReihe, 4] = S4;
                //oSheet.Cells[iReihe, 5] = S5;
                //oSheet.Cells[iReihe, 6] = S6;
                // dsizeX
                //oSheet.Cells[iReihe, 8] = S11;
                //oSheet.Cells[iReihe, 9] = S12;
                //oSheet.Cells[iReihe, 10] = S13;

                // Mittenprofil
                oSheet.Cells[iReihe, 1] = S31;
                oSheet.Cells[iReihe, 2] = S32;
                oSheet.Cells[iReihe, 3] = S33;


                // write bach the row counter in Excel
                sReihe = Convert.ToString(iReihe);
                oSheet.Cells[1, 1] = sReihe;
                // and save Excel
                oWB.Save();

            }
            // Quit after saving Excel
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(true, null, null);
                    oXL.Quit();
                }

            }
        }

        //#########################################################################################
        abListS SetupExcel(string ApplName, string DeptName,
                                string SetupDir, string SetupName,
                                string DebugDir, string DebugName,
                                string sheetDept, string sheetFolders, string sheetOrgans,
                                string sheetValues, string sheetTechn)

        {

            abListS abS = new abListS();

            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open(SetupDir + SetupName);
                oSheet = String.IsNullOrEmpty("ab_Values") ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["ab_Values"];

                //.................................................................................

                abS.Body = oSheet.Cells[5, 3].Value == null ? "-" : oSheet.Cells[5, 3].Value.ToString();

                abS.Hirn        = oSheet.Cells[6, 3].Value  == null ? "-" : oSheet.Cells[6, 3].Value.ToString();
                abS.Hirnstamm   = oSheet.Cells[7, 3].Value  == null ? "-" : oSheet.Cells[7, 3].Value.ToString();
                abS.Hippocampus = oSheet.Cells[8, 3].Value  == null ? "-" : oSheet.Cells[8, 3].Value.ToString();
                abS.Coclea      = oSheet.Cells[9, 3].Value  == null ? "-" : oSheet.Cells[9, 3].Value.ToString();
                abS.Chiasma     = oSheet.Cells[10, 3].Value == null ? "-" : oSheet.Cells[10, 3].Value.ToString();
                abS.Sehnerv     = oSheet.Cells[11, 3].Value == null ? "-" : oSheet.Cells[11, 3].Value.ToString();
                abS.Linse       = oSheet.Cells[12, 3].Value == null ? "-" : oSheet.Cells[12, 3].Value.ToString();

                abS.Parotis     = oSheet.Cells[13, 3].Value == null ? "-" : oSheet.Cells[13, 3].Value.ToString();
                abS.Submandi    = oSheet.Cells[14, 3].Value == null ? "-" : oSheet.Cells[14, 3].Value.ToString();
                abS.Mandibula   = oSheet.Cells[15, 3].Value == null ? "-" : oSheet.Cells[15, 3].Value.ToString();
                abS.Constrictor = oSheet.Cells[16, 3].Value == null ? "-" : oSheet.Cells[16, 3].Value.ToString();
                abS.Myelon      = oSheet.Cells[17, 3].Value == null ? "-" : oSheet.Cells[17, 3].Value.ToString();

                abS.Plexus      = oSheet.Cells[18, 3].Value == null ? "-" : oSheet.Cells[18, 3].Value.ToString();
                abS.Lunge       = oSheet.Cells[19, 3].Value == null ? "-" : oSheet.Cells[19, 3].Value.ToString();
                abS.Herz        = oSheet.Cells[20, 3].Value == null ? "-" : oSheet.Cells[20, 3].Value.ToString();
                abS.Riva        = oSheet.Cells[21, 3].Value == null ? "-" : oSheet.Cells[21, 3].Value.ToString();
                abS.Brust       = oSheet.Cells[22, 3].Value == null ? "-" : oSheet.Cells[22, 3].Value.ToString();

                abS.Oesophagus  = oSheet.Cells[23, 3].Value == null ? "-" : oSheet.Cells[23, 3].Value.ToString();
                abS.Leber       = oSheet.Cells[24, 3].Value == null ? "-" : oSheet.Cells[24, 3].Value.ToString();
                abS.Niere       = oSheet.Cells[25, 3].Value == null ? "-" : oSheet.Cells[25, 3].Value.ToString();
                abS.Darm        = oSheet.Cells[26, 3].Value == null ? "-" : oSheet.Cells[26, 3].Value.ToString();
                abS.Rektum      = oSheet.Cells[27, 3].Value == null ? "-" : oSheet.Cells[27, 3].Value.ToString();
                abS.Blase       = oSheet.Cells[28, 3].Value == null ? "-" : oSheet.Cells[28, 3].Value.ToString();
                abS.Gtv         = oSheet.Cells[29, 3].Value == null ? "-" : oSheet.Cells[29, 3].Value.ToString();
            }
            // Quit after saving Excel
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(true, null, null);
                    oXL.Quit();
                }

            }
            return abS;
        }
    }
}


 
