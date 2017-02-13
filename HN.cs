using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }
        public void Execute(ScriptContext context /*,System.Windows.Window window*/)
        {
            PlanSetup plan = context.PlanSetup;
            StructureSet ss = context.StructureSet;
            double presdose = plan.TotalPrescribedDose.Dose;
            string plansummary = string.Format("Course :\t{1}\nPlan :\t{0}\n\n",plan.Id.ToString(),plan.Course.Id.ToString());
            string msgfmt = plansummary;
            msgfmt += "Structure\t\tIntended Dose Index\tPlan\tJudge\n";
            msgfmt += Absmax("BRAIN_PRV", 7000, 1, ss, plan);
            msgfmt += Absmax("SPINAL_CORD_PRV",5000,1,ss,plan);
            msgfmt += Absmax("Spinal Canal", 5000, 1, ss, plan);
            msgfmt += Absmax("BRAIN_STEM_PRV", 5400, 1, ss, plan);
            msgfmt += Absmax("Brainstem", 5400, 1, ss, plan);
            msgfmt += Absmean("PAROTID_L",2600,1,ss,plan);
            msgfmt += Absmean("PAROTID_R", 2600, 1, ss, plan);
            msgfmt += "\n";
            msgfmt += Absmax("Chiasm", 5000, 1, ss, plan);
            msgfmt += Absmax("OPTIC_NRV_PRV", 5000, 1, ss, plan);
            msgfmt += Absmean("EAR_INN_L_PRV", 4500, 1, ss, plan);
            msgfmt += Absmean("EAR_INN_R_PRV", 4500, 1, ss, plan);
            msgfmt += Absmax("EYE", 4000, 1, ss, plan);
            msgfmt += Absmean("LENS_L", 600, 1, ss, plan);
            msgfmt += Absmean("LENS_R", 600, 1, ss, plan);
            msgfmt += Absmean("MUSCLE_CONST", 5400, 1, ss, plan);
            msgfmt += Absmean("LARYNX", 4500, 1, ss, plan);
            msgfmt += "\n";
            msgfmt += NTCPreturn("Brain", "Burman 1991, necrosis/infarct", ss, plan, 0.25, 0.15, 60);
            msgfmt += NTCPreturn("Brainstem", "Burman 1991, necrosis/infarct", ss, plan, 0.16, 0.14, 65);
            msgfmt += NTCPreturn("Spinal Canal", "Burman 1991, Myelitis/necrosis", ss, plan, 0.05, 0.175, 66.5);
            msgfmt += NTCPreturn("PAROTID_L", "Dijkema 2011, saliva flow < 25%", ss, plan, 1.13, 0.42, 39.4);
            msgfmt += NTCPreturn("PAROTID_R", "Dijkema 2011, saliva flow < 25%", ss, plan, 1.13, 0.42, 39.4);
            msgfmt += NTCPreturn("Chiasm", "Burman 1991, Blindness", ss, plan, 0.25, 0.14, 65);
            msgfmt += NTCPreturn("OPTIC_NRV_PRV", "Burman 1991, Blindness", ss, plan, 0.25, 0.14, 65);
            msgfmt += NTCPreturn("LENS_L", "Burman 1991, Cataract requiring intervention", ss, plan, 0.30, 0.27, 18);
            msgfmt += NTCPreturn("LENS_R", "Burman 1991, Cataract requiring intervention", ss, plan, 0.30, 0.27, 18);
            msgfmt += NTCPreturn("THYROID", "Bakhshandeh 2012, Hypothyroidism", ss, plan, 0.49, 0.24, 60);
            msgfmt += NTCPreturn("Tongue", "Musha, RTOG Grade over2", ss, plan, 0.03, 0.15, 72);
            msgfmt += NTCPreturn("Tongue", "Musha, RTOG Grade over1", ss, plan, 0.05, 0.23, 42.5);


            MessageBox.Show(msgfmt, "Head and Neck Dose Index");
            
        }
        public string NTCPreturn(string strnum, string gradename, StructureSet list, PlanSetup plan, double n, double m, double td50)
        {
            Structure retstructure = null;
            int bit = 0;
            double EUD = 0;
            double Ntcp = 0;
            string NTCPstring = "";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    bit = 1;
                    if (retstructure.IsEmpty == true) { bit = 2; break; }
                    DVHData dvh = plan.GetDVHCumulativeData(retstructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, 1);
                    var dvhdata = dvh.CurveData.ToList();
                    string volume = dvh.Volume.ToString();
                    string dvhstr = volume;
                    double fEUD = 0;
                    string dvhnums = dvhdata.Count.ToString();
                    for (int i = 0; i < dvhdata.Count - 1; i++)
                    {
                        fEUD += (dvhdata[i].Volume - dvhdata[i + 1].Volume) / 100 * Math.Pow(dvhdata[i].DoseValue.Dose / 100, 1 / n);
                    }
                    EUD = Math.Pow(fEUD, n);
                    double t = treturn(EUD, m, td50);
                    Ntcp = NTCP(t);
                    NTCPstring += string.Format("NTCP of {0}, {1} is {2}%\n", strnum, gradename, (Ntcp * 100).ToString("F2"));
                }
            }
            if (bit == 0)
            {
                NTCPstring += "";
                return NTCPstring;
            }
            else if (bit == 2) { NTCPstring += ""; return NTCPstring; }
            else { return NTCPstring; }
        }
        public double treturn(double EUD, double m, double td50)
        {
            double t = (EUD - td50) / m / td50;
            return t;
        }
        public double NTCP(double EUDt)
        {
            int steps = 1000;
            double intstart = -1000;
            double subtra = EUDt - intstart;//integration start from -1000
            double step = subtra / steps;
            double fntcp = 0;
            double mntcp = 0.0;
            for (int i = 1; i < steps - 1; i++)
            {
                mntcp += Math.Exp(-Math.Pow(intstart + step * i, 2) / 2);
            }
            fntcp = step / 2 * (Math.Exp(-Math.Pow(intstart, 2) / 2) + 2 * mntcp + Math.Exp(-Math.Pow(EUDt, 2) / 2));
            double NTCP = fntcp / Math.Sqrt(2 * Math.PI);
            return NTCP;
        }
        //public static DialogResult Show(string msgfmt)
        public string VaDp(string strnum, double doseindex, double criteria, int updown, StructureSet list, PlanSetup plan)//updown(1up,0down)
        {
            Structure retstructure = null;
            int bit = 0;
            double binwidth = 0.1;
            string doostr = "";
            string judge = "";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    if (retstructure.IsEmpty) { bit = 2; break; }
                    bit = 1;
                    DVHData dvhData = plan.GetDVHCumulativeData(retstructure, DoseValuePresentation.Absolute, VolumePresentation.Relative, binwidth);
                    DoseValue.DoseUnit doseUnit = dvhData.MaxDose.Unit;
                    Double Doo = plan.GetVolumeAtDose(retstructure, new DoseValue(doseindex, doseUnit), VolumePresentation.Relative);
                    doostr = Doo.ToString("F2");
                    if (updown == 1)//upper dose
                    {
                        if (Doo <= criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    else
                    {
                        if (Doo > criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    break;
                }
            }
            string doseindexreturn = "";
            string judgement = "";
            string retstring = "";
            if (bit == 1)
            {
                retstring = retstructure.Id; doseindexreturn = doostr; judgement = judge; retstring = string.Format("{0}\t\tV{1}Gy\t\t{2}%\t\t{3}%\t{4}\n", retstructure.Id, (doseindex / 100).ToString("F0"), criteria.ToString("F0"), doseindexreturn, judgement);
            }
            else if (bit == 2) { retstring = strnum+"\tis Empty structure"+"\n"; }
            else { retstring = "No Structure named " + strnum + "\n"; }
            return retstring;
        }
        public string Percmax(string strnum, double criteria, int updown, StructureSet list, PlanSetup plan)//updown(1up,0down)
        {
            Structure retstructure = null;
            int bit = 0;
            double binwidth = 0.1;
            string doostr = "";
            string judge = "";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    if (retstructure.IsEmpty) { bit = 2; break; }
                    bit = 1;
                    DVHData dvhData = plan.GetDVHCumulativeData(retstructure, DoseValuePresentation.Relative, VolumePresentation.AbsoluteCm3, binwidth);
                    DoseValue bbb = dvhData.MaxDose;
                    double Doo = bbb.Dose;
                    doostr = Doo.ToString("F2");
                    if (updown == 1)//upper dose
                    {
                        if (Doo <= criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    else
                    {
                        if (Doo > criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    break;
                }
            }
            string doseindexreturn = "";
            string judgement = "";
            string retstring = "";
            if (bit == 1)
            {
                retstring = retstructure.Id; doseindexreturn = doostr; judgement = judge; retstring = string.Format("{0}\t\tDmax\t\t{1}%\t\t{2}%\t{3}\n", retstructure.Id, criteria.ToString("F0"), doseindexreturn, judgement);
            }
            else if (bit == 2) { retstring = strnum + "\tis Empty structure" + "\n"; }
            else { retstring = "No Structure named " + strnum + "\n"; }
            return retstring;
        }
        public string Absmax(string strnum, double criteria, int updown, StructureSet list, PlanSetup plan)//updown(1up,0down)
        {
            Structure retstructure = null;
            int bit = 0;
            double binwidth = 0.1;
            string doostr = "";
            string judge = "";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    if (retstructure.IsEmpty) { bit = 2; break; }
                    bit = 1;
                    DVHData dvhData = plan.GetDVHCumulativeData(retstructure, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, binwidth);
                    DoseValue bbb = dvhData.MaxDose;
                    double Doo = bbb.Dose;
                    doostr = (Doo / 100).ToString("F2");
                    if (updown == 1)//upper dose
                    {
                        if (Doo <= criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    else
                    {
                        if (Doo > criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    break;
                }
            }
            string doseindexreturn = "";
            string judgement = "";
            string retstring = "";
            if (bit == 1)
            {
                retstring = retstructure.Id;
                doseindexreturn = doostr;
                judgement = judge;
                string ud = "";
                if (updown == 1) { ud = "<"; }
                else { ud = ">"; }
                retstring = string.Format("{0}\t\tDmax{4}{1}Gy\t{2}Gy\t{3}\n", retstructure.Id, (criteria / 100).ToString("F0"), doseindexreturn, judgement,ud);
            }
            else if (bit == 2) { retstring = strnum + "\tis Empty structure" + "\n"; }
            else { retstring = "No Structure named " + strnum + "\n"; }
            return retstring;
        }
        public string VpDp(string strnum, double doseindex, double criteria, int updown, StructureSet list, PlanSetup plan)//updown(1up,0down)
        {
            Structure retstructure = null;
            int bit = 0;
            double binwidth = 0.1;
            string doostr = "";
            string judge = "";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    if (retstructure.IsEmpty) { bit = 2; break; }
                    bit = 1;
                    DVHData dvhData = plan.GetDVHCumulativeData(retstructure, DoseValuePresentation.Relative, VolumePresentation.Relative, binwidth);
                    DoseValue.DoseUnit doseUnit = dvhData.MaxDose.Unit;
                    Double Doo = plan.GetVolumeAtDose(retstructure, new DoseValue(doseindex, doseUnit), VolumePresentation.Relative);
                    doostr = Doo.ToString("F2");
                    if (updown == 1)//upper dose
                    {
                        if (Doo <= criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    else
                    {
                        if (Doo >= criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    break;
                }
            }
            string doseindexreturn = "";
            string judgement = "";
            string retstring = "";
            if (bit == 1)
            {
                retstring = retstructure.Id; doseindexreturn = doostr; judgement = judge; retstring = string.Format("{0}\t\tV{1}%\t\t{2}%\t\t{3}%\t{4}\n", retstructure.Id, doseindex.ToString("F0"), criteria.ToString("F0"), doseindexreturn, judgement);
            }
            else if (bit == 2) { retstring = strnum + "\tis Empty structure" + "\n"; }
            else { retstring = "No Structure named " + strnum + "\n"; }
            return retstring;
        }
        public string VaDa(string strnum, double doseindex, double criteria, int updown, StructureSet list, PlanSetup plan)//updown(1up,0down)
        {
            Structure retstructure = null;
            int bit = 0;
            double binwidth = 0.1;
            string doostr = "";
            string judge = "";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    if (retstructure.IsEmpty) { bit = 2; break; }
                    bit = 1;
                    DVHData dvhData = plan.GetDVHCumulativeData(retstructure, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, binwidth);
                    DoseValue.DoseUnit doseUnit = dvhData.MaxDose.Unit;
                    Double Doo = plan.GetVolumeAtDose(retstructure, new DoseValue(doseindex, doseUnit), VolumePresentation.AbsoluteCm3);
                    doostr = Doo.ToString("F2");
                    if (updown == 1)//upper dose
                    {
                        if (Doo <= criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    else
                    {
                        if (Doo > criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    break;
                }
            }
            string doseindexreturn = "";
            string judgement = "";
            string retstring = "";
            if (bit == 1)
            {
                retstring = retstructure.Id; doseindexreturn = doostr; judgement = judge; retstring = string.Format("{0}\t\tV{1}Gy\t\t{2}cc\t\t{3}cc\t{4}\n", retstructure.Id, (doseindex / 100).ToString("F0"), criteria.ToString("F0"), doseindexreturn, judgement);
            }
            else if (bit == 2) { retstring = strnum + "\tis Empty structure" + "\n"; }
            else { retstring = "No Structure named " + strnum + "\n"; }
            return retstring;
        }
        public string Percmean(string strnum, double criteria, int updown, StructureSet list, PlanSetup plan)//updown(1up,0down)
        {
            Structure retstructure = null;
            int bit = 0;
            double binwidth = 0.1;
            string doostr = "";
            string judge = "";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    if (retstructure.IsEmpty) { bit = 2; break; }
                    bit = 1;
                    DVHData dvhData = plan.GetDVHCumulativeData(retstructure, DoseValuePresentation.Relative, VolumePresentation.AbsoluteCm3, binwidth);
                    DoseValue bbb = dvhData.MeanDose;
                    double Doo = bbb.Dose;
                    doostr = Doo.ToString("F2");
                    if (updown == 1)//upper dose
                    {
                        if (Doo <= criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    else
                    {
                        if (Doo > criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    break;
                }
            }
            string doseindexreturn = "";
            string judgement = "";
            string retstring = "";
            if (bit == 1)
            {
                retstring = retstructure.Id; doseindexreturn = doostr; judgement = judge; retstring = string.Format("{0}\t\tDmean\t\t{1}%\t\t{2}%\t{3}\n", retstructure.Id, criteria.ToString("F0"), doseindexreturn, judgement);
            }
            else if (bit == 2) { retstring = strnum + "\tis Empty structure" + "\n"; }
            else { retstring = "No Structure named " + strnum + "\n"; }
            return retstring;
        }
        public string Absmean(string strnum, double criteria, int updown, StructureSet list, PlanSetup plan)//updown(1up,0down)
        {
            Structure retstructure = null;
            int bit = 0;
            double binwidth = 0.1;
            string doostr = "";
            string judge = "";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    if (retstructure.IsEmpty) { bit = 2; break; }
                    bit = 1;
                    DVHData dvhData = plan.GetDVHCumulativeData(retstructure, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, binwidth);
                    DoseValue bbb = dvhData.MeanDose;
                    double Doo = bbb.Dose;
                    doostr = (Doo / 100).ToString("F2");
                    if (updown == 1)//upper dose
                    {
                        if (Doo <= criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    else
                    {
                        if (Doo > criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    break;
                }
            }
            string doseindexreturn = "";
            string judgement = "";
            string retstring = "";
            if (bit == 1)
            {
                retstring = retstructure.Id;
                doseindexreturn = doostr;
                judgement = judge;
                string ud = "0";
                if (updown == 1) { ud = "<"; }
                else { ud = ">"; }
                retstring = string.Format("{0}\t\tDmean{4}{1}Gy\t{2}Gy\t{3}\n", retstructure.Id, (criteria / 100).ToString("F0"), doseindexreturn, judgement,ud);
            }
            else if (bit == 2) { retstring = strnum + "\tis Empty structure" + "\n"; }
            else { retstring = "No Structure named " + strnum + "\n"; }
            return retstring;
        }
        static string Prostatevol(StructureSet list)
        {
            Structure retstructure = null;
            string retstring = "No Prostate exist!";
            string strnum = "Prostate";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    double strvol = retstructure.Volume;
                    string fracstr = strvol.ToString("F2");
                    retstring = "Volume of Prostate " + fracstr + "cc\n";
                }
            }
            return retstring;
        }
        public string DaVp(string strnum, double doseindex, double criteria, int updown, StructureSet list, PlanSetup plan)//updown(1up,0down)
        {
            Structure retstructure = null;
            int bit = 0;
            string doostr = "";
            string judge = "";
            foreach (Structure scan in list.Structures)
            {
                if (scan.Id == strnum)
                {
                    retstructure = scan;
                    if (retstructure.IsEmpty) { bit = 2; break; }
                    bit = 1;
                    DoseValue Do = plan.GetDoseAtVolume(retstructure, doseindex, VolumePresentation.Relative, DoseValuePresentation.Absolute);
                    Double Doo = Do.Dose / 100;
                    doostr = Doo.ToString("F2");
                    if (updown == 1)//upper dose
                    {
                        if (Doo <= criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    else
                    {
                        if (Doo > criteria) { judge = "o"; }
                        else { judge = "x"; }
                    }
                    break;
                }
            }
            string doseindexreturn = "";
            string judgement = "";
            string retstring = "";
            if (bit == 1)
            {
                retstring = retstructure.Id; doseindexreturn = doostr; judgement = judge; retstring = string.Format("{0}\t\tD{1}%\t\t{2}Gy\t\t{3}Gy\t{4}\n", retstructure.Id, doseindex.ToString("F0"), (criteria / 100).ToString("F0"), doseindexreturn, judgement);
            }
            else if (bit == 2) { retstring = strnum + "\tis Empty structure" + "\n"; }
            else { retstring = "No Structure named " + strnum + "\n"; }
            return retstring;
        }
    }
}

