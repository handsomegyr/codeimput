#region Using directives

using System.Collections.Generic;
using System.IO;
using System.Text;
using System;
using System.Windows.Forms;
using MongoDB.Driver;
using MongoDB.Bson;

#endregion

namespace ExcelApplication1
{
    public class Worker
    {

        public static bool merge(string selectedPath, List<ExcelFile> excelFileList,Settings settings ,System.ComponentModel.BackgroundWorker backgroundWorker)
        {
            // Start the search for primes and wait.
            UTF8Encoding utf8 = new UTF8Encoding(false);
            var targetPath = settings.TargetPath;
            var writer = new StreamWriter(targetPath, false, utf8);

            try
            {
                int i = 0;
                foreach (var f in excelFileList)
                {
                    i++;
                    if (backgroundWorker.CancellationPending)
                    {
                        // Return without doing any more work.
                        throw new Exception("用户取消了操作");
                    }

                    if (backgroundWorker.WorkerReportsProgress)
                    {
                        backgroundWorker.ReportProgress(i);
                    }

                    if (!f.IsSelected) continue;
                    f.Status = "处理中";
                    var stream = new FileStream(f.Path, FileMode.Open);
                    StreamReader txtReader = new StreamReader(stream, utf8);
                    try
                    {

                        string nextLine;                        
                        while ((nextLine = txtReader.ReadLine()) != null)
                        {
                            //虚拟卡编号,虚拟卡密码,开始有效期,截止有效期,活动名称,奖品名称,是否使用,活动编号
                            var _id =ObjectId.GenerateNewId().ToString();// 序号
                            var code = nextLine.Trim();
                            var pwd = "";
                            var start_time = settings.StartTime;
                            var end_time = settings.EndTime;
                            var prize_id = settings.PrizeId;
                            var is_used = settings.IsUsed ? 1 : 0;

                            DateTime now = DateTime.Now;

                            //按yyyy-MM-dd HH:mm:ss格式输出s
                            var __CREATE_TIME__ = now.ToString("yyyy-MM-dd HH:mm:ss");
                            var __MODIFY_TIME__ = now.ToString("yyyy-MM-dd HH:mm:ss");

                            if (!string.IsNullOrEmpty(code))
                            {
                                writer.WriteLine("insert  into `iprize_code`(`_id`,`prize_id`,`code`,`pwd`,`is_used`,`start_time`,`end_time`,`__CREATE_TIME__`,`__MODIFY_TIME__`,`__REMOVED__`) values ('{0}','{1}','{2}','{3}',{4},'{5}','{6}','{7}','{8}',0);", _id, prize_id, code, pwd, is_used, start_time, end_time, __CREATE_TIME__, __MODIFY_TIME__);
                            }
                        }

                        f.Status = "处理完成";
                    }
                    catch (Exception e1)
                    {
                        f.Status = "处理失败";
                        throw e1;
                    }
                    finally
                    {;
                        txtReader.Close();
                    }
                }
                
                return true;
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
                throw e;
                //return false;
            }
            finally
            {
                writer.Flush();
                writer.Close();
            }
        }

        
    }
}
