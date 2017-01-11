using System;
using System.Collections.Generic;
using System.Text;
using JHEvaluation.ScoreCalculation.ScoreStruct;
using SCSemsScore = JHEvaluation.ScoreCalculation.ScoreStruct.SemesterScore;
using System.Linq;
using System.Windows.Forms;

namespace JHEvaluation.ScoreCalculation.BigFunction
{
    class SemesterScoreCalculator
    {
        private List<StudentScore> Students { get; set; }

        private UniqueSet<string> Filter { get; set; }

        Dictionary<string, string> Dic_StudentAttendScore = new Dictionary<string, string>();

        Dictionary<string, List<string>> errCheck = new Dictionary<string, List<string>>();

        public SemesterScoreCalculator(List<StudentScore> students, IEnumerable<string> filter)
        {
            Students = students;
            Filter = new UniqueSet<string>(filter);
        }

        /// <summary>
        /// 計算科目成績
        /// </summary>
        public void CalculateSubjectScore()
        {
            foreach (StudentScore student in Students)
            {
                if (student.CalculationRule == null) continue; //沒有成績計算規則就不計算。

                SCSemsScore Sems = student.SemestersScore[SemesterData.Empty];

                //對全部的學期科目成績作一次擇優
                foreach (string subject in Sems.Subject)
                {
                    if (Sems.Subject.Contains(subject))
                    {
                        SemesterSubjectScore sss = Sems.Subject[subject];

                        //沒有原始成績就將既有的成績填入
                        if (!sss.ScoreOrigin.HasValue)
                            sss.ScoreOrigin = sss.Value;

                        //有原始或補考成績才做擇優
                        if (sss.ScoreOrigin.HasValue || sss.ScoreMakeup.HasValue)
                            sss.BetterScoreSelection();
                    }
                }



                // 2016/7/13 穎驊註解，原本為整理重覆科目使用，後來功能被另外取代掉，又保留Dic_StudentAttendScore 的話，
                // 如果再 //處理修課課程 foreach (string subject in student.AttendScore)，不使用in student.AttendScore 而是用 in Dic_StudentAttendScore
                // 則會出現 下一個學生使用上一個學生的修課紀錄，假如上一個學生有著下一個學生沒有的修課內容，就會造成錯誤


                //foreach (string subject in student.AttendScore)
                //{
            
                //    if (!Dic_StudentAttendScore.ContainsKey(subject))
                //    {
                //        Dic_StudentAttendScore.Add(subject, subject);

                //    }
                //    else {

                //        //MessageBox.Show("科目:" + subject + "在課程科目名稱有重覆，" + "如繼續匯入將導致學期科目成績遺漏誤植，請確認您在" + subject + "這一科的課程資料無誤");
                //    }
                                                                               
                //}



                //處理修課課程
                foreach (string subject in student.AttendScore)
                {
                    if (!IsValidItem(subject)) continue; //慮掉不算的科目。

                    AttendScore attend = student.AttendScore[subject];
                    SemesterSubjectScore SemsSubj = null;
                    
                    if (Sems.Subject.Contains(subject))
                        SemsSubj = Sems.Subject[subject];
                    else
                        SemsSubj = new SemesterSubjectScore();

                    if (attend.ToSemester) //列入學期成績。
                    {
                        //分數、權數都有資料才進行計算科目成績。
                        if (attend.Value.HasValue && attend.Weight.HasValue && attend.Period.HasValue)
                        {
                            decimal score = student.CalculationRule.ParseSubjectScore(attend.Value.Value); //進位處理。
                            //SemsSubj.Value = score;
                            SemsSubj.Weight = attend.Weight.Value;
                            SemsSubj.Period = attend.Period.Value;
                            SemsSubj.Effort = attend.Effort;
                            SemsSubj.Text = attend.Text;
                            SemsSubj.Domain = attend.Domain;

                            //填到原始成績
                            SemsSubj.ScoreOrigin = score;
                            //擇優成績
                            SemsSubj.BetterScoreSelection();
                        }
                        else //資料不合理，保持原來的分數狀態。
                            continue;

                        if (!Sems.Subject.Contains(subject))
                            Sems.Subject.Add(subject, SemsSubj);
                    }
                    else
                    {
                        //不計入學期成績，就將其刪除掉。
                        if (Sems.Subject.Contains(subject))
                            Sems.Subject.Remove(subject);
                    }
                }
            }
        }

        /// <summary>
        /// 計算領域成績
        /// </summary>
        /// <param name="defaultRule"></param>
        //public void CalculateDomainScore(ScoreCalculator defaultRule,bool clearDomainScore)
        public void CalculateDomainScore(ScoreCalculator defaultRule,DomainScoreSetting setting)
        {
            EffortMap effortmap = new EffortMap(); //努力程度對照表。

            // 高雄領域轉換用
            string khDomain = "語文";

            foreach (StudentScore student in Students)
            {
                SemesterScore semsscore = student.SemestersScore[SemesterData.Empty];
                
                SemesterDomainScoreCollection dscores = semsscore.Domain;



                //if (dscores.Contains("國語文") && student.SemestersScore[SemesterData.Empty].Domain["國語文"].ScoreMakeup!=null) {

                //    dscores["國語文"].ScoreMakeup = student.SemestersScore[SemesterData.Empty].Domain["國語文"].ScoreMakeup;
                
                //}


                SemesterSubjectScoreCollection jscores = semsscore.Subject;
                ScoreCalculator rule = student.CalculationRule;

                if (rule == null)
                    rule = defaultRule;

                //各領域分數加總。
                Dictionary<string, decimal> domainTotal = new Dictionary<string, decimal>();
                //各領域原始分數加總。
                Dictionary<string, decimal> domainOriginTotal = new Dictionary<string, decimal>();
                //各領域權重加總。
                Dictionary<string, decimal> domainWeight = new Dictionary<string, decimal>();
                //各領域節數加總。
                Dictionary<string, decimal> domainPeriod = new Dictionary<string, decimal>();
                //文字評量串接。
                Dictionary<string, string> domainText = new Dictionary<string, string>();

                //各領域補考分數加總。
                Dictionary<string, decimal> domain_MakeUpScore_Total = new Dictionary<string, decimal>();

                //被限制不能超過60分的領域(該領域的科目有補考成績)
                List<string> LimitedDomains = new List<string>();

                //該領域的科目有補考成績清單
                List<string> Have_MakeUpScore_Domains = new List<string>();

                // 只計算領域成績，不計算科目
                bool OnlyCalcDomainScore = false;

                // 檢查科目成績是否有成績，當有成績且成績的科目領域名稱非空白才計算
                int CheckSubjDomainisNotNullCot = 0;

                // 檢查是否有非科目領域是空白的科目
                foreach (string str in jscores)
                {
                    SemesterSubjectScore objSubj = jscores[str];
                    string strDomainName = objSubj.Domain.Trim();
                    if (!string.IsNullOrEmpty(strDomainName))
                        CheckSubjDomainisNotNullCot++;

                    //該領域的科目有補考成績將被加入
                    if (objSubj.ScoreMakeup.HasValue && !LimitedDomains.Contains(strDomainName))
                    {
                        LimitedDomains.Add(strDomainName);

                        if (strDomainName == "國語文" || strDomainName == "英語")
                        {
                            LimitedDomains.Add(khDomain);
                        }    
                    }

                    //該領域的科目有補考成績將被加入
                    if (objSubj.ScoreMakeup.HasValue && !Have_MakeUpScore_Domains.Contains(strDomainName))
                    {                  
                        Have_MakeUpScore_Domains.Add(strDomainName);

                        if (strDomainName == "國語文" || strDomainName == "英語")
                        {
                            Have_MakeUpScore_Domains.Add(khDomain);
                        }                        
                    }                                     
                }

                // 當沒有科目成績或科目成績內領域沒有非空白，只計算領域成績。
                if (jscores.Count == 0 || CheckSubjDomainisNotNullCot == 0)
                    OnlyCalcDomainScore = true;
                else
                    OnlyCalcDomainScore = false;


                if (OnlyCalcDomainScore == false)
                {
                    // 從科目計算到領域

                    #region 總計各領域的總分、權重、節數。
                    foreach (string strSubj in jscores)
                    {
                        SemesterSubjectScore objSubj = jscores[strSubj];
                        string strDomain = objSubj.Domain.Injection();

                        //不計算的領域就不算。
                        if (!IsValidItem(strDomain)) continue;

                        if (objSubj.Value.HasValue && objSubj.Weight.HasValue && objSubj.Period.HasValue)
                        {
                            if (!objSubj.ScoreOrigin.HasValue)
                                objSubj.ScoreOrigin = objSubj.Value;

                            // 針對高雄處理
                            if (strDomain == "國語文" || strDomain == "英語")
                            {
                                if (!domainTotal.ContainsKey(khDomain))
                                {
                                    domainTotal.Add(khDomain, 0);
                                    domainOriginTotal.Add(khDomain, 0);
                                    domainWeight.Add(khDomain, 0);
                                    domainPeriod.Add(khDomain, 0);
                                    // 2016/2/26 經高雄繼斌與蔡主任討論，語文領域文字描述不需要儲存
                                    //domainText.Add(khDomain, string.Empty);

                                    //領域補考成績
                                    domain_MakeUpScore_Total.Add(khDomain, 0);
                                }
                                
                                domainTotal[khDomain] += objSubj.Value.Value * objSubj.Weight.Value;
                                //科目的原始成績加總
                                domainOriginTotal[khDomain] += objSubj.ScoreOrigin.Value * objSubj.Weight.Value;

                                domainWeight[khDomain] += objSubj.Weight.Value;
                                domainPeriod[khDomain] += objSubj.Period.Value;
                                // 2016/2/26 經高雄繼斌與蔡主任討論，語文領域文字描述不需要儲存
                                domainText[khDomain] = "";//+= GetDomainSubjectText(strSubj, objSubj.Text);

                                //領域補考成績總和 ，算法為 先用各科目的"成績" 加權計算 ， 之後會再判斷， 此成績總和 是否含有"補考成績" 是否為"補考成績總和"
                                domain_MakeUpScore_Total[khDomain] += objSubj.Value.Value * objSubj.Weight.Value;
                            }

                            if (!domainTotal.ContainsKey(strDomain))
                            {
                                domainTotal.Add(strDomain, 0);
                                domainOriginTotal.Add(strDomain, 0);
                                domainWeight.Add(strDomain, 0);
                                domainPeriod.Add(strDomain, 0);
                                domainText.Add(strDomain, string.Empty);

                                //領域補考成績
                                domain_MakeUpScore_Total.Add(strDomain, 0);
                            }

                            domainTotal[strDomain] += objSubj.Value.Value * objSubj.Weight.Value;
                            //科目的原始成績加總
                            domainOriginTotal[strDomain] += objSubj.ScoreOrigin.Value * objSubj.Weight.Value;

                            domainWeight[strDomain] += objSubj.Weight.Value;
                            domainPeriod[strDomain] += objSubj.Period.Value;
                            domainText[strDomain] += GetDomainSubjectText(strSubj, objSubj.Text);

                            //領域補考成績總和 ，算法為 先用各科目的"成績" 加權計算 ， 之後會再判斷， 此成績總和 是否含有"補考成績" 是否為"補考成績總和"
                            domain_MakeUpScore_Total[strDomain] += objSubj.Value.Value * objSubj.Weight.Value;
                        }
                    }
                    #endregion

                    #region 計算各領域加權平均。

                    // 高雄市特有語文領域成績
                    decimal sTotal = 0, sTotalOrigin = 0, sWeight = 0, sPeriod = 0, sWeightAvg = 0, sWeightOriginAvg = 0;


                    //2016/5/20(蔡英文就職) 穎驊新增高雄市特有語文領域 努力程度計算方式

                    decimal Chinese_Effort = 0, English_Effort = 0, Chinese_Weight = 0, English_Weight = 0, s_Language_Weight = 0;


                    string sText = "";

                    /* (2014/11/18 補考調整)調整為保留領域成績中的資訊，但是移除多的領域成績項目。 */
                    // 從科目算過來清空原來領域成績，以科目成績的領域為主
                    //dscores.Clear();

                    foreach (string strDomain in domainTotal.Keys)
                    {

                        if (Program.Mode == ModuleMode.KaoHsiung)
                        {
                            if (strDomain == "國語文" || strDomain == "英語")
                            {
                                sTotal += domainTotal[strDomain];
                                sTotalOrigin += domainOriginTotal[strDomain];
                                sWeight += domainWeight[strDomain];
                                sPeriod += domainPeriod[strDomain];
                                // 2016/2/26 經高雄繼斌與蔡主任討論，語文領域文字描述不需要儲存
                                sText = "";// string.Join(";", domainText[strDomain].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries));                           


                                if (domainWeight.ContainsKey("國語文") && domainWeight["國語文"] != null)
                                {

                                    Chinese_Weight = domainWeight["國語文"];
                                }

                                if (domainWeight.ContainsKey("英語") && domainWeight["英語"] != null)
                                {
                                    English_Weight = domainWeight["英語"];
                                }
                                s_Language_Weight = Chinese_Weight + English_Weight;

                            }
                        }


                        decimal total = domainTotal[strDomain];
                        decimal totalOrigin = domainOriginTotal[strDomain];

                        decimal weight = domainWeight[strDomain];
                        decimal period = domainPeriod[strDomain];
                        string text = "";

                        // 補考總分
                        decimal makeup_score_total = domain_MakeUpScore_Total[strDomain];

                        

                        if (domainText.ContainsKey(strDomain))
                            text = string.Join(";", domainText[strDomain].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries));

                        if (weight <= 0) continue; //沒有權重就不計算，保留原來的成績。

                        decimal weightAvg = rule.ParseDomainScore(total / weight);
                        decimal weightOriginAvg = rule.ParseDomainScore(totalOrigin / weight);

                        // 補考總分平均
                        decimal? makeup_score_total_Avg = rule.ParseDomainScore(makeup_score_total / weight);

                        //將成績更新回學生。
                        SemesterDomainScore dscore = null;
                        if (dscores.Contains(strDomain))
                            dscore = dscores[strDomain];
                        else
                        {
                            dscore = new SemesterDomainScore();
                            dscores.Add(strDomain, dscore);
                        }

                        //先將算好的成績帶入領域成績,後面的擇優判斷才不會有問題
                        dscore.Value = weightAvg;
                        dscore.ScoreOrigin = weightOriginAvg;
                        dscore.Weight = weight;
                        dscore.Period = period;
                        dscore.Text = text;
                        dscore.Effort = effortmap.GetCodeByScore(weightAvg);

                        dscore.ScoreMakeup = Have_MakeUpScore_Domains.Contains(strDomain) ? makeup_score_total_Avg : null;

                        if (strDomain == "國語文")
                        {
                            Chinese_Effort = effortmap.GetCodeByScore(weightAvg);
                        }

                        if (strDomain == "英語")
                        {
                            English_Effort = effortmap.GetCodeByScore(weightAvg);

                        }

                        //後面會有一段code執行全領域的擇優判斷,此段已不需要
                        //若有補考成績就進行擇優,否則就將科目的擇優平均帶入領域成績
                        //if (dscore.ScoreOrigin.HasValue || dscore.ScoreMakeup.HasValue)
                        //dscore.BetterScoreSelection(); //進行成績擇優。
                        //else
                        //dscore.Value = weightAvg;
                    }

                    //清除不應該存在領域成績
                    //if (clearDomainScore)
                    if (setting.DomainScoreClear)
                    {
                        foreach (var domainName in dscores.ToArray())
                        {
                            //如果新計算的領域成績中不包含在原領域清單中，就移除他。
                            if (!domainTotal.ContainsKey(domainName))
                                dscores.Remove(domainName);
                        }
                    }

                    //高雄市特有領域成績
                    if (Program.Mode == ModuleMode.KaoHsiung)
                    {
                        if (sWeight > 0)
                        {
                            SemesterDomainScore dsSS = null;
                            if (dscores.Contains("語文"))
                                dsSS = dscores["語文"];
                            else
                            {
                                dsSS = new SemesterDomainScore();
                                dscores.Add("語文", dsSS);
                            }

                            //sWeightAvg = rule.ParseDomainScore(sTotal / sWeight);

                            sWeightOriginAvg = rule.ParseDomainScore(sTotalOrigin / sWeight);

                            //dsSS.Value = sWeightAvg;

                            dsSS.ScoreOrigin = sWeightOriginAvg;
                            dsSS.Weight = sWeight;
                            dsSS.Period = sPeriod;
                            dsSS.Text = sText;


                            //dsSS.Effort = effortmap.GetCodeByScore(sWeightAvg);

                            //2016/5/20 穎驊更正語文領域努力程度正確計算顯示
                            decimal Language_Effort = ((Chinese_Effort * Chinese_Weight) + (English_Effort * English_Weight)) / s_Language_Weight;

                            // 將原本為decimal Language_Effort 型別 四捨五入後 轉型成int
                            int Language_Effort_int = Convert.ToInt32(decimal.Round(Language_Effort, 0, MidpointRounding.AwayFromZero));

                            dsSS.Effort = Language_Effort_int;

                            // 檢查國語文與英語補考成績，如果有加權平均填入語文
                            bool hasmmScore = false;
                            decimal scSum = 0;
                            decimal mmScore = 0;
                            if (dscores.Contains("國語文"))
                            {
                                if (dscores["國語文"].ScoreMakeup.HasValue && dscores["國語文"].Weight.HasValue)
                                {
                                    mmScore += dscores["國語文"].ScoreMakeup.Value * dscores["國語文"].Weight.Value;
                                    scSum += dscores["國語文"].Value.Value * dscores["國語文"].Weight.Value;
                                    hasmmScore = true;
                                }
                                else
                                {
                                    if (dscores["國語文"].Value.HasValue)
                                    {
                                        mmScore += dscores["國語文"].Value.Value * dscores["國語文"].Weight.Value;
                                        scSum += dscores["國語文"].Value.Value * dscores["國語文"].Weight.Value;
                                    }
                                }
                            }

                            if (dscores.Contains("英語"))
                            {
                                if (dscores["英語"].ScoreMakeup.HasValue && dscores["英語"].Weight.HasValue)
                                {
                                    mmScore += dscores["英語"].ScoreMakeup.Value * dscores["英語"].Weight.Value;
                                    scSum += dscores["英語"].Value.Value * dscores["英語"].Weight.Value;
                                    hasmmScore = true;
                                }
                                else
                                {
                                    if (dscores["英語"].Value.HasValue)
                                    {
                                        mmScore += dscores["英語"].Value.Value * dscores["英語"].Weight.Value;
                                        scSum += dscores["英語"].Value.Value * dscores["英語"].Weight.Value;
                                    }
                                }
                            }

                            if (hasmmScore)
                                dsSS.ScoreMakeup = rule.ParseDomainScore(mmScore / sWeight);

                            dsSS.Value = rule.ParseDomainScore(scSum / sWeight);

                        }
                    }


                    #endregion
                }

                # region 處理沒有學期科目成績狀況

                // 2016/5/24 穎驊處理  OnlyCalcDomainScore = true狀況(原本並沒有else處理)，也就是該學生該學期沒有學期科目成績，
                //只有領域成績，此現象可能發生在學期中轉入的轉學生上

                //2016/7/11 穎驊改動，因應克服需求，原本要更正OnlyCalcDomainScore = true狀況的計算方式，但跟恩正、國志討論後，
                //如果學校方在計算時遇到沒有學習科目成績的學生，就不再另外計算、提醒，避免以後麻煩


                else
                {
                    //// 高雄市特有語文領域成績
                    //decimal sTotal = 0, sTotalOrigin = 0, sWeight = 0, sPeriod = 0, sWeightAvg = 0, sWeightOriginAvg = 0;

                    ////2016/5/20(蔡英文就職) 穎驊新增高雄市特有語文領域 努力程度計算方式

                    //decimal Chinese_Effort = 0, English_Effort = 0, Chinese_Weight = 0, English_Weight = 0, s_Language_Weight = 0;

                    //string sText = "";


                    ////高雄市特有領域成績
                    //if (Program.Mode == ModuleMode.KaoHsiung)
                    //{
                    //    if (dscores.Contains("國語文") && dscores["國語文"].Value.HasValue && dscores["國語文"].ScoreOrigin.HasValue && dscores["國語文"].Weight.HasValue
                    //        && dscores["國語文"].Period.HasValue)
                    //    {


                    //        sTotal += (decimal)dscores["國語文"].Value;
                    //        sTotalOrigin += (decimal)dscores["國語文"].ScoreOrigin * (decimal)dscores["國語文"].Weight;
                    //        sWeight += (decimal)dscores["國語文"].Weight;
                    //        sPeriod += (decimal)dscores["國語文"].Period;

                    //        Chinese_Effort = (decimal)dscores["國語文"].Effort;
                    //        Chinese_Weight = (decimal)dscores["國語文"].Weight;

                    //    }

                    //    else {



                    //        if (!errCheck.ContainsKey(student.Name +student.Id))
                    //        {

                    //            errCheck.Add(student.Name + student.Id, new List<string>());
                    //        }
                    //        errCheck[student.Name + student.Id].Add("請確認學生:" + student.Name + " 在所計算學期內" + "學期領域:國語文的每一項數值是否有輸入不正確、遺漏");
                   
                    //    }

                    //    if (dscores.Contains("英語") && dscores["英語"].Value.HasValue && dscores["英語"].ScoreOrigin.HasValue && dscores["英語"].Weight.HasValue
                    //         && dscores["英語"].Period.HasValue)
                    //    {
                    //        sTotal += (decimal)dscores["英語"].Value;
                    //        sTotalOrigin += (decimal)dscores["英語"].ScoreOrigin * (decimal)dscores["英語"].Weight;
                    //        sWeight += (decimal)dscores["英語"].Weight;
                    //        sPeriod += (decimal)dscores["英語"].Period;


                    //        English_Effort = (decimal)dscores["英語"].Effort;
                    //        English_Weight = (decimal)dscores["英語"].Weight;
                    //    }

                    //    else
                    //    {
                    //        if (!errCheck.ContainsKey(student.Name + student.Id))
                    //        {

                    //            errCheck.Add(student.Name + student.Id, new List<string>());
                    //        }
                    //        errCheck[student.Name + student.Id].Add("請確認學生:" + student.Name + " 在所計算學期內" + "學期領域:英語的每一項數值是否有輸入不正確、遺漏");
                    //    }


                    //    if (sWeight > 0)
                    //    {
                    //        SemesterDomainScore dsSS = null;
                    //        if (dscores.Contains("語文"))
                    //            dsSS = dscores["語文"];
                    //        else
                    //        {
                    //            dsSS = new SemesterDomainScore();
                    //            dscores.Add("語文", dsSS);
                    //        }


                    //        //sWeightAvg = rule.ParseDomainScore(sTotal / sWeight);

                    //        sWeightOriginAvg = rule.ParseDomainScore(sTotalOrigin / sWeight);

                    //        s_Language_Weight = English_Weight + Chinese_Weight;
                    //        //dsSS.Value = sWeightAvg;


                    //        dsSS.ScoreOrigin = sWeightOriginAvg;
                    //        dsSS.Weight = sWeight;
                    //        dsSS.Period = sPeriod;
                    //        dsSS.Text = sText;
                    //        //dsSS.Effort = effortmap.GetCodeByScore(sWeightAvg);


                    //        //2016/5/20 穎驊更正語文領域努力程度正確計算顯示
                    //        decimal Language_Effort = ((Chinese_Effort * Chinese_Weight) + (English_Effort * English_Weight)) / s_Language_Weight;

                    //        // 將原本為decimal Language_Effort 型別 四捨五入後 轉型成int
                    //        int Language_Effort_int = Convert.ToInt32(decimal.Round(Language_Effort, 0, MidpointRounding.AwayFromZero));

                    //        dsSS.Effort = Language_Effort_int;

                    //        // 檢查國語文與英語補考成績，如果有加權平均填入語文
                    //        bool hasmmScore = false;
                    //        decimal scSum = 0;
                    //        decimal mmScore = 0;
                    //        if (dscores.Contains("國語文"))
                    //        {
                    //            if (dscores["國語文"].ScoreMakeup.HasValue && dscores["國語文"].Weight.HasValue)
                    //            {
                    //                mmScore += dscores["國語文"].ScoreMakeup.Value * dscores["國語文"].Weight.Value;
                    //                hasmmScore = true;
                    //            }
                    //            else
                    //            {


                    //                if (dscores["國語文"].Value.HasValue)
                    //                {
                    //                    mmScore += dscores["國語文"].Value.Value * dscores["國語文"].Weight.Value;
                    //                    scSum += dscores["國語文"].Value.Value * dscores["國語文"].Weight.Value;
                    //                }
                    //            }
                    //        }

                    //        if (dscores.Contains("英語"))
                    //        {
                    //            if (dscores["英語"].ScoreMakeup.HasValue && dscores["英語"].Weight.HasValue)
                    //            {
                    //                mmScore += dscores["英語"].ScoreMakeup.Value * dscores["英語"].Weight.Value;
                    //                hasmmScore = true;
                    //            }
                    //            else
                    //            {
                    //                if (dscores["英語"].Value.HasValue)
                    //                {
                    //                    mmScore += dscores["英語"].Value.Value * dscores["英語"].Weight.Value;
                    //                    scSum += dscores["英語"].Value.Value * dscores["英語"].Weight.Value;
                    //                }
                    //            }
                    //        }

                    //        if (hasmmScore)
                    //            dsSS.ScoreMakeup = rule.ParseDomainScore(mmScore / sWeight);

                    //        dsSS.Value = rule.ParseDomainScore(scSum / sWeight);

                    //    }
                    //}

                }

                #endregion


                //這段會對全部領域做一次擇優
                foreach (var domain in dscores.ToArray())
                {
                    SemesterDomainScore objDomain = dscores[domain];

                    //沒有原始成績就將既有的成績填入
                    //if (!objDomain.ScoreOrigin.HasValue)
                       // objDomain.ScoreOrigin = objDomain.Value;

                    if (objDomain.ScoreOrigin.HasValue || objDomain.ScoreMakeup.HasValue)
                        objDomain.BetterScoreSelection();

                    //領域成績被限制為不能超過60分
                    // 2017/1/11 穎驊筆記，因應 恩正與高雄小組重新討論後，有了以下 成績新判斷法，
                    // 若領域補考 沒有大於原始分數 則領域成績分數為 領域原始成績 (以本程式碼， 上面會先有 BetterScoreSelection() 擇優， 下面第一個if 判斷 是否 要直接用原始成績)
                    // (我知道 這樣子做 很像繞圈圈走回頭冗路，為什麼 不直接 在BetterScoreSelection() 裡更動就好?)
                    // (因為我是接手 維護此CODE ，先前的CODE 架構已經太複雜， 我都盡量不會再更動，避免有意外發生。)
                    // 若領域補考 大於原始分數 且又小於60分 則領域成績分數  為 領域補考成績
                    // 若領域補考 大於原始分數 且又大於60分 則領域成績分數 以60分 為上限

                    //詳情可看本專案 Resource 資料匣 內流程圖圖檔 : 補考科目成績計算流程圖  TXT 檔: 補考科目成績計算流程文字
                    // 或是 高雄項目 [08-07][06] 補考科目成績影響領域成績計算結果問題 的討論

                    if (setting.DomainScoreLimit)
                    {
                        if (LimitedDomains.Contains(domain) && objDomain.ScoreOrigin > 60 )
                        {
                            objDomain.Value = objDomain.ScoreOrigin;
                        }

                        if (LimitedDomains.Contains(domain) && 60 > objDomain.ScoreMakeup && objDomain.ScoreMakeup > objDomain.ScoreOrigin)
                        {
                            objDomain.Value = objDomain.ScoreMakeup;
                        }

                        if (LimitedDomains.Contains(domain) &&  objDomain.ScoreMakeup >60  &&  60 >objDomain.ScoreOrigin )
                        {
                            objDomain.Value = 60;
                        }                        
                    }
                }

     

                //計算課程學習成績。
                ScoreResult result = CalcDomainWeightAvgScore(dscores, new UniqueSet<string>());
                if (result.Score.HasValue)
                    semsscore.CourseLearnScore = rule.ParseLearnDomainScore(result.Score.Value);
                if (result.ScoreOrigin.HasValue)
                    semsscore.CourseLearnScoreOrigin = rule.ParseLearnDomainScore(result.ScoreOrigin.Value);

                //計算學習領域成績。
                result = CalcDomainWeightAvgScore(dscores, Util.VariableDomains);
                if (result.Score.HasValue)
                    semsscore.LearnDomainScore = rule.ParseLearnDomainScore(result.Score.Value);
                if (result.ScoreOrigin.HasValue)
                    semsscore.LearnDomainScoreOrigin = rule.ParseLearnDomainScore(result.ScoreOrigin.Value);
            }


            //if (errCheck.Count > 0)
            //{
            //    StringBuilder sb = new StringBuilder();
                
            //    foreach (var student in errCheck.Keys)
            //    {
            //        foreach (var err in errCheck[student])
            //        {
            //            sb.AppendLine(string.Format("{0}",err));
            //        }
            //    }
            //    MessageBox.Show(sb.ToString());
            //}


        }

        /// <summary>
        /// 單純從學期科目成績的文字描述加總到學期領域成績的文字描述
        /// </summary>
        /// <param name="defaultRule"></param>
        public void SumDomainTextScore()
        {
            foreach (StudentScore student in Students)
            {
                SemesterScore semsscore = student.SemestersScore[SemesterData.Empty];
                SemesterDomainScoreCollection dscores = semsscore.Domain;
                SemesterSubjectScoreCollection jscores = semsscore.Subject;

                Dictionary<string, string> domainText = new Dictionary<string, string>();

                #region 總計各領域的總分、權重、節數。
                foreach (string strSubj in jscores)
                {
                    SemesterSubjectScore objSubj = jscores[strSubj];
                    string strDomain = objSubj.Domain.Injection();

                    //不計算的領域就不算。
                    if (!IsValidItem(strDomain)) continue;

                    if (objSubj.Value.HasValue && objSubj.Weight.HasValue && objSubj.Period.HasValue)
                    {
                        if (!domainText.ContainsKey(strDomain))
                        {
                            domainText.Add(strDomain, string.Empty);
                        }

                        domainText[strDomain] += GetDomainSubjectText(strSubj, objSubj.Text);
                    }
                }
                #endregion

                #region 計算各領域加權平均。
                foreach (string strDomain in domainText.Keys)
                {
                    string text = string.Join(";", domainText[strDomain].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries));

                    //將成績更新回學生。
                    SemesterDomainScore dscore = null;
                    if (dscores.Contains(strDomain))
                        dscore = dscores[strDomain];
                    else
                    {
                        dscore = new SemesterDomainScore();
                        dscores.Add(strDomain, dscore);
                    }

                    dscore.Text = text;
                }
                #endregion
            }
        }

        /// <summary>
        /// 計算學習領域成績
        /// </summary>
        /// <param name="defaultRule"></param>
        public void CalculateLearningDomainScore(ScoreCalculator defaultRule)
        {
            foreach (StudentScore student in Students)
            {
                SemesterScore semsscore = student.SemestersScore[SemesterData.Empty];
                SemesterDomainScoreCollection dscores = semsscore.Domain;
                ScoreCalculator rule = student.CalculationRule;

                if (rule == null)
                    rule = defaultRule;

                #region 計算各領域加權平均。
                //計算學習領域成績。
                //decimal? result = CalcDomainWeightAvgScore(dscores, Util.VariableDomains);
                ScoreResult result = CalcDomainWeightAvgScore(dscores, Util.VariableDomains); ;
                if (result.Score.HasValue)
                    semsscore.LearnDomainScore = rule.ParseLearnDomainScore(result.Score.Value);

                if (result.ScoreOrigin.HasValue)
                    semsscore.LearnDomainScoreOrigin = rule.ParseLearnDomainScore(result.ScoreOrigin.Value);
                #endregion

            }
        }

        private static ScoreResult CalcDomainWeightAvgScore(SemesterDomainScoreCollection dscores, UniqueSet<string> excludeItem)
        {
            decimal TotalScore = 0, TotalScoreOrigin = 0, TotalWeight = 0;
            foreach (string strDomain in dscores)
            {
                // 高雄語文只顯示不計算
                if (Program.Mode == ModuleMode.KaoHsiung)
                    if (strDomain == "語文")
                        continue;

                if (excludeItem.Contains(strDomain.Trim())) continue;

                SemesterDomainScore dscore = dscores[strDomain];
                if (dscore.Value.HasValue && dscore.Weight.HasValue) //dscore.Value 是原來的結構。
                {
                    TotalScore += (dscore.Value.Value * dscore.Weight.Value); //擇優成績。

                    if (dscore.ScoreOrigin.HasValue)
                        TotalScoreOrigin += (dscore.ScoreOrigin.Value * dscore.Weight.Value); //原始成績。
                    else
                        TotalScoreOrigin += (dscore.Value.Value * dscore.Weight.Value); //將擇優當成原始。

                    TotalWeight += dscore.Weight.Value; //比重不會因為哪種成績而不同。
                }
            }

            if (TotalWeight <= 0) return new ScoreResult(); //沒有成績。

            ScoreResult sr = new ScoreResult();

            sr.Score = TotalScore / TotalWeight;
            sr.ScoreOrigin = TotalScoreOrigin / TotalWeight;

            return sr;
        }

        private class ScoreResult
        {
            /// <summary>
            /// 成績，一般是擇優後的成績，大部份的報表都會使用此成績。
            /// </summary>
            public decimal? Score { get; set; }

            /// <summary>
            /// 原始成績，最原始的成績。
            /// </summary>
            public decimal? ScoreOrigin { get; set; }
        }

        private string GetDomainSubjectText(string subjName, string subjText)
        {
            if (!string.IsNullOrEmpty(subjText))
            {
                return string.Format("({0}){1};", subjName, subjText);
            }
            else
                return string.Empty;
        }

        private bool IsValidItem(string item)
        {
            if (Filter.Count <= 0) return true;

            return (Filter.Contains(item));
        }
    }
}
