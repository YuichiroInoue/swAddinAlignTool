using System;
using System.Collections.Generic;
using System.IO;
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
using Microsoft.Win32;
using SldWorks;
using SwConst;

namespace swAddinAlignTool
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        static SldWorks.SldWorks swApp;
        static ModelDoc2 swModel;
        static string DIR = System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory)+@"\";
        static string PARTNAME = "";
        static string INFOPATH =DIR+ @"size_info.csv";

        public MainWindow()
        {
            InitializeComponent();
            path.Text = DIR;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Solidworksのプロセス起動
            swApp = new SldWorks.SldWorks();
            //画面を表示する falseなら裏で動作
            swApp.Visible = true;


            string fileName = null;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "step(*.step)|*.step;*.stp|parasolid(*.x_t)|*.x_t";
            if(ofd.ShowDialog()==true)
            {
                fileName = ofd.FileName;
            }else
            {
                return;
            }
            PARTNAME = fileName.Substring(fileName.LastIndexOf(@"\")+1);

            //新規モデルの読み込み
            bool bRet = false;
            string strArg = null;
            ImportStepData importData = default(ImportStepData);
            int Err = 0;

            //fileName = DIR+"v1_4.step";

            importData = swApp.GetImportFileData(fileName);
            swModel = swApp.LoadFile4(fileName, strArg, importData,ref Err);


        }

        public double GetAngle(double[] v1,double[] v2)
        {
            double retAngle;
            double cos, a1, b1, a2, b2;
            a1 = v1[0];
            a2 = v1[1];
            b1 = v2[0];
            b2 = v2[1];
            cos = (a1 * b1 + a2 * b2) / (Math.Sqrt(Math.Pow(a1,2) + Math.Pow(a2,2))
                * (Math.Sqrt(Math.Pow(b1,2)+ Math.Pow(b2,2))));
            if (double.IsNaN(cos)) 
            {
                retAngle = 0;
                return retAngle;
            }
            else if(cos==0)
            {
                retAngle = 90;
                return retAngle;
            }
            double retRad = Math.Acos(cos);
            retAngle = Math.Round((180 * retRad) / Math.PI,6);
            return retAngle;
        }

        /// <summary>
        /// z align
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (swApp == null) { return; }
            PartDoc swPart = (PartDoc)swModel;

            //モデルを選択して移動/コピーフィーチャーを挿入
            SelectionMgr swSelMgr = default(SelectionMgr);
            SelectData swSelData = default(SelectData);
            Body2 swBody = default(Body2);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swSelData = (SelectData)swSelMgr.CreateSelectData();


            FeatureManager ftMgr = default(FeatureManager);
            ftMgr = swModel.FeatureManager;
            double transX, transY, transZ, transD, rotPX, rotPY, rotPZ, rotAX, rotAY, rotAZ;
            transX = 0;
            transY = 0;
            transZ = 0;
            transD = 0;
            rotPX = 0;
            rotPY = 0;
            rotPZ = 0;
            rotAX = 0;
            rotAY = 0;
            rotAZ = 0;
            int numCopies = 1;

            //選択面の法線ベクトルを取得
            Face2 swFace = default(Face2);
            double[] faceNormalVector = new double[3];

            //ここで面を選択しておくと、選択面がXY平面と水平になるように回転する。
            try
            {
                swFace = swSelMgr.GetSelectedObject6(1, -1);
            }catch
            {
                MessageBox.Show("select z face");
                return;
            }
            if (swFace==null)
            {
                MessageBox.Show("select z face");
                return;
            }
            faceNormalVector = swFace.Normal;

            double[] xyV = new double[2];
            xyV[0] = Math.Round(faceNormalVector[0], 6);
            xyV[1] = Math.Round(faceNormalVector[1], 6);
            double[] yzV = new double[2];
            yzV[0] = Math.Round(faceNormalVector[1], 6);
            yzV[1] = Math.Round(faceNormalVector[2], 6);
            double[] yUnitV = new double[2] { 0, 1 };
            double[] zUnitV = new double[2] { 0, 1 };

            double xyAngle = GetAngle(xyV, yUnitV);
            double yzAngle = GetAngle(yzV, zUnitV);

            rotAZ = xyAngle * Math.PI / 180;
            rotAX = yzAngle * Math.PI / 180;

            object[] bodyArr = null;
            bodyArr = swPart.GetBodies2((int)swBodyType_e.swAllBodies, false);

            swBody = (Body2)bodyArr[0];
            swSelData.Mark = 1;
            bool bRet = swBody.Select2(true, swSelData);

            Feature ftMoveCopy = ftMgr.InsertMoveCopyBody2(transX, transY, transZ, transD,
    rotPX, rotPY, rotPZ, rotAZ, rotAY, rotAX, false, numCopies);

        }

        /// <summary>
        /// xy align
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (swApp == null) { return; }
            PartDoc swPart = (PartDoc)swModel;

            //モデルを選択して移動/コピーフィーチャーを挿入
            SelectionMgr swSelMgr = default(SelectionMgr);
            SelectData swSelData = default(SelectData);
            Body2 swBody = default(Body2);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swSelData = (SelectData)swSelMgr.CreateSelectData();


            FeatureManager ftMgr = default(FeatureManager);
            ftMgr = swModel.FeatureManager;
            double transX, transY, transZ, transD, rotPX, rotPY, rotPZ, rotAX, rotAY, rotAZ;
            transX = 0;
            transY = 0;
            transZ = 0;
            transD = 0;
            rotPX = 0;
            rotPY = 0;
            rotPZ = 0;
            rotAX = 0;
            rotAY = 0;
            rotAZ = 0;
            int numCopies = 1;

            //選択面の法線ベクトルを取得
            Face2 swFace = default(Face2);
            double[] faceNormalVector = new double[3];

            //ここで面を選択しておくと、選択面がXY平面と水平になるように回転する。
            try
            {
                swFace = swSelMgr.GetSelectedObject6(1, -1);
            }catch
            {
                MessageBox.Show("select y face");
                return;
            }
            if (swFace==null)
            {
                MessageBox.Show("select y face");
                return;
            }
            faceNormalVector = swFace.Normal;

            double zComp = Math.Round(faceNormalVector[2], 6);
            if(zComp!=0)
            {
                //error
                MessageBox.Show("Select faces parallel to Z axis");
                return;
            }
            double[] xyV = new double[2];
            xyV[0] = Math.Round(faceNormalVector[0], 6);
            xyV[1] = Math.Round(faceNormalVector[1], 6);
            double[] yUnitV = new double[2] { 0, 1 };

            double xyAngle = GetAngle(xyV, yUnitV);

            rotAZ = xyAngle * Math.PI / 180;

            object[] bodyArr = null;
            bodyArr = swPart.GetBodies2((int)swBodyType_e.swAllBodies, false);

            swBody = (Body2)bodyArr[0];
            swSelData.Mark = 1;
            bool bRet = swBody.Select2(true, swSelData);

            Feature ftMoveCopy = ftMgr.InsertMoveCopyBody2(transX, transY, transZ, transD,
    rotPX, rotPY, rotPZ, rotAZ, rotAY, rotAX, false, numCopies);


        }

        /// <summary>
        /// center align
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            if (swApp == null) { return; }
            double[] bodyBox = new double[6];

            PartDoc swPart = (PartDoc)swModel;
            object[] bodyArr = null;
            bodyArr = swPart.GetBodies2((int)swBodyType_e.swAllBodies, false);
            Body2 swBody = (Body2)bodyArr[0];

            bodyBox = swBody.GetBodyBox();
            double x1, y1, z1, x2, y2, z2;
            x1 = bodyBox[0];
            y1 = bodyBox[1];
            z1 = bodyBox[2];
            x2 = bodyBox[3];
            y2 = bodyBox[4];
            z2 = bodyBox[5];

            double xCenter, yCenter, xMove, yMove ,zMove;
            xCenter = (x1 + x2) / 2;
            yCenter = (y1 + y2) / 2;
            xMove = xCenter*-1;
            yMove = yCenter*-1;
            if(z1>z2)
            {
                zMove = z1;
            }else
            {
                zMove = -z2;
            }

            SelectionMgr swSelMgr = default(SelectionMgr);
            SelectData swSelData = default(SelectData);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;
            swSelData = (SelectData)swSelMgr.CreateSelectData();
            swSelData.Mark = 1;
            bool bRet = swBody.Select2(true, swSelData);

            FeatureManager ftMgr = default(FeatureManager);
            ftMgr = swModel.FeatureManager;
            Feature ftMoveCopy = ftMgr.InsertMoveCopyBody2(xMove, yMove, zMove, 0, 0, 0, 0, 0, 0, 0, false, 1);

            double xSize, ySize, zSize;
            xSize = Math.Round(Math.Abs(x1*1000 - x2*1000), 6);
            ySize = Math.Round(Math.Abs(y1 * 1000 - y2 * 1000), 6);
            zSize = Math.Round(Math.Abs(z1 * 1000 - z2 * 1000), 6);
            using(System.IO.StreamWriter sw = new System.IO.StreamWriter(INFOPATH,true))
            {
                string strInfo = PARTNAME + ","
                    + xSize.ToString() + "," + ySize.ToString() + "," + zSize.ToString();
                sw.WriteLine(strInfo);
            }
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            if (swApp == null) { return; }
            ModelDocExtension swModelExt = swModel.Extension;
            string exportStepFileName = PARTNAME.Remove(PARTNAME.IndexOf(".")) + ".step";
            string fileName = path.Text+ @"set\"+exportStepFileName;
            Directory.CreateDirectory(path.Text + "set");

            ExportPdfData exportPdf = default;
            int Err, Warn;
            Err = 0;
            Warn = 0;
            bool bRet;

            swModel.ClearSelection2(true);
            bRet = swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swStepAP, 214);

            bRet = swModelExt.SaveAs(fileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 0, exportPdf, Err, Warn);
            if (bRet)
            {
                MessageBox.Show("Completed successfully");
            }
            else
            {
                MessageBox.Show("export incomplete");
            }
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            if (swApp != null)
            {
                swApp.ExitApp();
                swApp = null;
            }
        }
    }
}
