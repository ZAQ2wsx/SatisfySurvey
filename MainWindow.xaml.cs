using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

using Path = System.IO.Path;

using static SatisfySurvey.WordHelper;

namespace SatisfySurvey
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        bool excuteOnce;

        public MainWindow()
        {
            InitializeComponent();
            excuteOnce = false;
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                textBox.Text = args[1];
                button_Click(this, null);
            }
        }

        /// <summary>
        /// 统计按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Click(object sender, RoutedEventArgs e)
        {
            if (excuteOnce)
                return;

            excuteOnce = true;

            InitHelper();

            prgBar.Visibility = Visibility.Visible;
            string[] names = Directory.GetFiles(textBox.Text);
            
            nums = names.Length;

            new Thread(() => { SomeTask(names, ChangeLabel); }).Start();
            
        }

        void SomeTask(string[] names,Action<int> action)
        {
            
            for (int i = 0; i < names.Length; i++)
            {
                Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action<int>(ChangeLabel),i );
                if (names[i].Contains(".doc") && !names[i].Contains("~"))
                {
                    try
                    {
                        WordHelper.DealWord(names[i]);
                    }
                    catch (Exception e)
                    {
                        WordHelper.errorList.Add("name: " + names[i] + "  exception:" + e);
                    }
                }
            }
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action<int>(ChangeLabel), names.Length);
        }
        
        int nums;
        private void ChangeLabel(int index)
        {
            prgBar.Value = (index) * 100/nums;
            Console.WriteLine(prgBar.Value + " value!");
            if (index == nums)
            {
                string outp = WordHelper.Show();

                if (WordHelper.errorList?.Count > 0)
                {
                    File.WriteAllText(Path.Combine(textBox.Text, "错误.txt"), outp);

                    output.Content = "统计出错，详情已输出错误日志";

                }
                else
                {
                    File.WriteAllText(Path.Combine(textBox.Text, "统计结果.txt"), outp);

                    output.Content = "已输出统计结果 至 选择文件夹, 点击 浏览 查看";
                }
                Process[] p = Process.GetProcessesByName("wps");
                Console.WriteLine(p.Length + " pL");
                foreach (var pp in p)
                {
                    pp.Kill();
                }
                Process[] px = Process.GetProcessesByName("WINWORD");
                Console.WriteLine(px.Length + " pL");
                foreach (var pp in px)
                {
                    pp.Kill();
                }


                if (checkBox.IsChecked == true)
                {
                    button2_Click(this, null);
                    System.Environment.Exit(0);
                }
            }
        }

        /// <summary>
        /// 更改目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            if (excuteOnce)
                return;

            FolderBrowserDialog fialog = new FolderBrowserDialog();
            if (fialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox.Text = fialog.SelectedPath;
            }
        }

        /// <summary>
        /// 打开输出目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("explorer", textBox.Text);
        }
        
    }


}
