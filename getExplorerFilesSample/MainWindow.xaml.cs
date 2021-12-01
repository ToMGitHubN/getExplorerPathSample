using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Shell32;
using SHDocVw;

namespace getExplorerFilesSample
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // エクスプローラ(Explorer.EXE)で開いているフォルダのパスを取得する
            // http://tinqwill.blog59.fc2.com/blog-entry-84.html#comment11


            //リストボックスをクリアしておく
            listBox1.Items.Clear();
            //
            //COMのShellクラス作成
            Shell shell = new Shell();
            //IEとエクスプローラの一覧を取得
            ShellWindows win = shell.Windows();
            //列挙
            foreach (IWebBrowser2 web in win)
            {
                //エクスプローラのみ(IEを除外)
                if (System.IO.Path.GetFileName(web.FullName).ToUpper() == "EXPLORER.EXE")
                {
                    //リストに追加
                    listBox1.Items.Add(web.LocationURL + " : " + web.LocationName);
                }
            }
        }
    }
}
