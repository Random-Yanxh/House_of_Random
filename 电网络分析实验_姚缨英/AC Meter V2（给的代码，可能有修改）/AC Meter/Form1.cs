using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SeeSharpTools.JY.ArrayUtility;
using SeeSharpTools.JY.GUI;
using JYUSB61902;
using SeeSharpTools.JY.Statistics;
using SeeSharpTools.JY.DSP.FilteringMCR;
using SeeSharpTools.JY.DSP.SoundVibration;
using MicroLibrary;

namespace AC_Meter
{
    public partial class Form1 : Form
    {
       
        private JYUSB61902AITask aitask;                      //采集卡任务初始化       
        private readonly MicroLibrary.MicroTimer microTimer;  //定时器初始化

        double[,] readvalue;     // 存放采集卡原始数据
        double[,] display;       // 用于波形显示 
        double[] V1instant,V2instant,V3instant;       // 存放电压测量原始值  
        double[] Vabinstant, Vbcinstant, Vcainstant;  // 线电压瞬时值
        double[] Iinstant;       // 存放电流测量原始值
        double[] Iqinstant;      // 存放电流移相90度的值
        double[] Pinstant;       // 存放功率瞬时值     
        double[] Qinstant;       // 存放无功功率瞬时值
        double scale1 = 1, scale2 = 1, scale3 = 1;
        double V1rms, V2rms, V3rms;
        double Vabrms, Vbcrms, Vcarms;
        double Irms, Vrms,P, S, Q, cos_fi, fi;  //计算值，第一种算法
        double P2,S2,Q2, cos_fi2, fi2;          //计算值，第二种算法           
        double Phase1, Phase2, Phase3, Phase4;

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        int index = 0;
 
        int[] peak1 = new int[5];
        int[] peak2 = new int[5];
        int[] peak3 = new int[5];
        int[] peak4 = new int[5];
        bool Digital_Filter;     //是否使用数字滤波             
        int Ichannel = 0;
        double factor = 2 * 3.1415 * 80 / 50000;
        bool Ismeasuring = false;//开始测量的标志位        
        public Vector Vch1, Vch2, Vch3, Ich4;
        ToneAnalysisResult Vch1tone, Vch2tone, Vch3tone, Ich4tone;

        public class Vector
        {
            public double Amp = 0;
            public double Angle = 0;
        }        

        public Form1()
        { 
            InitializeComponent();              //系统初始化
            aitask = new JYUSB61902AITask(0);   //定义采集任务对应的采集卡地址            

            aitask.AddChannel(0, -10, 10, AITerminal.RSE);  //添加模拟量采集通道0，电压范围-10，10，方式为单端
            aitask.AddChannel(1, -10, 10, AITerminal.RSE);  //添加模拟量采集通道1，电压范围-10，10，方式为单端
            aitask.AddChannel(2, -10, 10, AITerminal.RSE);  //添加模拟量采集通道0，电压范围-10，10，方式为单端
            aitask.AddChannel(3, -10, 10, AITerminal.RSE);  //添加模拟量采集通道1，电压范围-10，10，方式为单端

            aitask.Mode = AIMode.Continuous;  //设定采集方式为连续采集
            aitask.SampleRate = 50000;       //设定每个通道的采样率50k

            //初始化存放采集原始数据的数组
            readvalue = new double[5000, 4]; 
            display = new double[4, 5000];                
            V1instant = new double[5000];       
            V2instant = new double[5000];  
            V3instant = new double[5000];
            Vabinstant = new double[5000];
            Vbcinstant = new double[5000];
            Vcainstant = new double[5000];
            Iinstant = new double[5000];
            Iqinstant = new double[5000];
            Pinstant = new double[5000];
            Qinstant = new double[5000];

            Digital_Filter = false;          //滤波器是否开启        

            microTimer = new MicroLibrary.MicroTimer(); //定义定时器及定时处理
            microTimer.MicroTimerElapsed +=
                new MicroLibrary.MicroTimer.MicroTimerElapsedEventHandler(OnTimedEvent);

            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 3;

            Vch1 = new Vector();
            Vch2 = new Vector();
            Vch3 = new Vector();
            Ich4 = new Vector();

        }
            
        private void easyButton1_Click(object sender, EventArgs e)
        {
            if(!Ismeasuring)
            {
                Ismeasuring = true;           //测量标志位置位
                aitask.Start();               //启动采集
                microTimer.Interval = 100000; //定时器间隔，单位微妙
                microTimer.Start();           //开启定时器              
                easyButton1.Text = "停止测量";
                index = 0;
            }
            else
            {
                Ismeasuring = false;               
                easyButton1.Enabled = true;
                easyButton1.Text = "开始测量";
            }                      
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
                scale1 = 1;
            else if (comboBox1.SelectedIndex == 1)
                scale1 = 10;
            else if (comboBox1.SelectedIndex == 2)
                scale1 = 100;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 0)
                scale2 = 1;
            else if (comboBox2.SelectedIndex == 1)
                scale2 = 10;
            else if (comboBox2.SelectedIndex == 2)
                scale2 = 100;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex == 0)
                scale3 = 1;
            else if (comboBox3.SelectedIndex == 1)
                scale3 = 10;
            else if (comboBox3.SelectedIndex == 2)
                scale3 = 100;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.SelectedIndex == 0)
                Ichannel = 1;
            else if (comboBox4.SelectedIndex == 1)
                Ichannel = 2;
            else if (comboBox4.SelectedIndex == 2)
                Ichannel = 3;
            else if (comboBox4.SelectedIndex == 3)
                Ichannel = 0;
        }

        private void easyButton2_Click(object sender, EventArgs e)
        {
            if (!Digital_Filter)
            {
                
                Digital_Filter = true;
                easyButton2.Text = "关闭滤波";
            }
            else
            {
                Digital_Filter = false;
                easyButton2.Text = "开启滤波";
            }
        }


        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            
            int x, y;
            double Vx, Vy, angle, amp;

            Graphics VecDiagram = e.Graphics;
            Pen pblack = new Pen(Color.Black, 1);
            Pen pyellow = new Pen(Color.Yellow, 3);
            pyellow.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
            Pen pred = new Pen(Color.Red, 3);
            pred.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
            Pen pblue = new Pen(Color.Blue, 3);
            pblue.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
            Pen pgreen = new Pen(Color.Green, 3);
            pgreen.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;

            VecDiagram.DrawEllipse(pblack, 25, 25, 300, 300); //画圈            
            VecDiagram.DrawEllipse(pblack, 50, 50, 250, 250); //画圈        
            VecDiagram.DrawEllipse(pblack, 75, 75, 200, 200); //画圈        
            VecDiagram.DrawEllipse(pblack, 100, 100, 150, 150); //画圈        
            VecDiagram.DrawEllipse(pblack, 125, 125, 100, 100); //画圈        
            VecDiagram.DrawEllipse(pblack, 150, 150, 50, 50); //画圈        
            VecDiagram.DrawLine(pblack, 0, 175, 350, 175);   //画个十字
            VecDiagram.DrawLine(pblack, 175, 0, 175, 350);
                       
            amp = Vch1.Amp*0.65;
            angle = Vch1.Angle+90;
            Vx = 175 + amp * Math.Sin(angle / 180 * Math.PI);
            Vy = 175 + amp * Math.Cos(angle / 180 * Math.PI);
            x = (int)Vx;
            y = (int)Vy;
            VecDiagram.DrawLine(pred, 175, 175, x, y);

            amp = Vch2.Amp*0.65;
            angle = -Vch2.Angle+90;
            Vx = 175 + amp * Math.Sin(angle / 180 * Math.PI);
            Vy = 175 + amp * Math.Cos(angle / 180 * Math.PI);
            x = (int)Vx;
            y = (int)Vy;
            VecDiagram.DrawLine(pgreen, 175, 175, x, y);

            amp = Vch3.Amp*0.65;
            angle = -Vch3.Angle+90;
            Vx = 175 + amp * Math.Sin(angle / 180 * Math.PI);
            Vy = 175 + amp * Math.Cos(angle / 180 * Math.PI);
            x = (int)Vx;
            y = (int)Vy;
            VecDiagram.DrawLine(pyellow, 175, 175, x, y);

            amp = Ich4.Amp*500;
            angle = -Ich4.Angle+90;
            Vx = 175 + amp * Math.Sin(angle / 180 * Math.PI);
            Vy = 175 + amp * Math.Cos(angle / 180 * Math.PI);
            x = (int)Vx;
            y = (int)Vy;
            VecDiagram.DrawLine(pblue, 175, 175, x, y);

        }

        private double[] Digitalfilter(double[] sourcedata)
        {
            int i;
            int length;
            length = sourcedata.Length;

            double[] filtereddata = new double[length];
            filtereddata[0] = factor * sourcedata[0];
            i = 1;
            while (i< length)
            {
                filtereddata[i] = (1-factor) * filtereddata[i - 1] + factor * sourcedata[i];
                i++;
            }
            return filtereddata;   


        }
        
        private void OnTimedEvent(object sender,MicroLibrary.MicroTimerEventArgs timerEventArgs)
        {
            /*定时中断处理。定时器设定为0.1s，每次触发中断。中断处理进行原始数据采集，
            *分离获取电压、电流、功率等瞬时测量值，计算处理得到有效值、功率值、功率因数
            *值等。将得到的波形进行绘图，并在用户界面显示计算结果。*/
            index++;
            if (!Ismeasuring)
            {
                aitask.Stop();      //停止采集
                microTimer.Stop();  //停止定时器                
            }
            else
            {
                if (aitask.AvailableSamples > 5000)
                //判断是否采集卡缓存区有足够数据。采样率100k，0.1s中断，所以每次读取数据是10k。
                {
                    aitask.ReadData(ref readvalue);  //采集获取缓存区的原始测量数据
                    int i = 0;
                    while (i < 5000)
                    {
                        V1instant[i] = readvalue[i, 0] * scale1;
                        V2instant[i] = readvalue[i, 1] * scale2;
                        V3instant[i] = readvalue[i, 2] * scale3;
                        Iinstant[i] = readvalue[i, 3];
                        Iqinstant[i] = Iinstant[(i + 250) % 5000];   //电流移相90度的瞬时值。采样率50k,1周期对应1k数组元素。90度相位对应250
                        i++;
                    }

                    if (Digital_Filter)
                    {
                        /*数字滤波器滤波*/
                        Iinstant = Digitalfilter(Iinstant);
                        V1instant = Digitalfilter(V1instant);
                        V2instant = Digitalfilter(V2instant);
                        V3instant = Digitalfilter(V3instant);
                    }

                    i = 0;
                    while(i<5000)
                    {
                        display[0, i] = V1instant[i];
                        display[1, i] = V2instant[i];
                        display[2, i] = V3instant[i];
                        display[3, i] = Iinstant[i];
              
                        Vabinstant[i] = V1instant[i] - V2instant[i];
                        Vbcinstant[i] = V2instant[i] - V3instant[i];
                        Vcainstant[i] = V3instant[i] - V1instant[i];
                        i++;
                    }

                    V1rms = Statistics.RMS(V1instant);
                    V2rms = Statistics.RMS(V2instant);
                    V3rms = Statistics.RMS(V3instant);
                    Irms = Statistics.RMS(Iinstant);
                    Vabrms = Statistics.RMS(Vabinstant);
                    Vbcrms = Statistics.RMS(Vbcinstant);
                    Vcarms = Statistics.RMS(Vcainstant);

                    //Vch1tone= HarmonicAnalyzer.ToneAnalysis(V1instant,1,10,true);
                    //Vch2tone = HarmonicAnalyzer.ToneAnalysis(V2instant, 1, 10, true);
                    //Vch3tone = HarmonicAnalyzer.ToneAnalysis(V3instant, 1, 10, true);
                    //Ich4tone = HarmonicAnalyzer.ToneAnalysis(Iinstant, 1, 10, true);
                    

                    i = 0;
                    if (Ichannel==1)
                    {                        
                        while (i < 5000)
                        {
                            Pinstant[i] = V1instant[i] * Iinstant[i];  //功率计算
                            Qinstant[i] = V1instant[i] * Iqinstant[i]; //无功功率计算
                            i++;
                        }
                        Vrms = V1rms;
                    }
                    else if (Ichannel == 2)
                    {
                        while (i < 5000)
                        {
                            Pinstant[i] = V2instant[i] * Iinstant[i];  //功率计算
                            Qinstant[i] = V2instant[i] * Iqinstant[i]; //无功功率计算
                            i++;
                        }
                        Vrms = V2rms;
                    }
                    else if (Ichannel == 3)
                    {
                        while (i < 5000)
                        {
                            Pinstant[i] = V3instant[i] * Iinstant[i];  //功率计算
                            Qinstant[i] = V3instant[i] * Iqinstant[i]; //无功功率计算
                            i++;
                        }
                        Vrms = V3rms;
                    }

                    /*计算有功功率、无功功率、功率因数等等*/
                    P = Statistics.Mean(Pinstant);     // 调用库函数，计算真有效值。5个周期       

                    /* 第一种算法 */
                    S = Vrms * Irms;
                    Q = Math.Sqrt(S * S - P * P);
                    cos_fi = P / S;
                    fi = Math.Acos(cos_fi) * 180 / 3.1415;

                    /* 第二种算法 */
                    Q2 = Statistics.Mean(Qinstant);
                    S2 = Math.Sqrt(Q2 * Q2 + P * P);
                    cos_fi2 = P / S2;
                    fi2 = Math.Acos(cos_fi2) * 180 / 3.1415;

                    /*计算相位及矢量角度*/
                    i = 1;
                    int j1 = 0, j2 = 0, j3 = 0, j4 = 0;
                    while(i<4999)
                    {
                        if ((V1instant[i]>=0) && (V1instant[i - 1]<0) && (V1instant[i+1]>0))
                        {
                            if (j1 <= 4)
                            {
                                peak1[j1] = i;
                                j1++;
                            }                            
                        }
                        if ((V2instant[i] >= 0) && (V2instant[i - 1] < 0) && (V2instant[i + 1] > 0))
                        {
                            if (j2 <= 4)
                            {
                                peak2[j2] = i;
                                j2++;
                            }              
                        }
                        if ((V3instant[i] >=0) && (V3instant[i - 1] < 0) && (V3instant[i + 1] > 0))
                        {
                            if (j3 <= 4)
                            {
                                peak3[j3] = i;
                                j3++;
                            }        
                        }
                        if ((Iinstant[i] >= 0) && (Iinstant[i - 1] < 0) && (Iinstant[i + 1] > 0))
                        {
                            if (j4 <= 4)
                            {
                                peak4[j4] = i;
                                j4++;
                            }                                                    
                        }
                        i++;
                    }

                    i = 0;
                    Phase1 = 0;
                    Phase2 = 0;
                    Phase3 = 0;
                    Phase4 = 0;
                    while(i<3)
                    {
                        Phase1 = Phase1 + peak1[i] - i*1000;
                        Phase2 = Phase2 + peak2[i] -i* 1000;
                        Phase3 = Phase3 + peak3[i] -i* 1000;
                        Phase4 = Phase4 + peak4[i] -i* 1000;
                        i++;
                    }
                    Phase1 = Phase1 * 360 / 3000;
                    Phase2 = Phase2 * 360 / 3000;
                    Phase3 = Phase3 * 360 / 3000;
                    Phase4 = Phase4 * 360 / 3000;

                    Vch1.Amp = V1rms;
                    Vch1.Angle = 0;
                    Vch2.Amp = V2rms;
                    Vch2.Angle = Phase2-Phase1;
                    if (Vch2.Angle < -360)
                        Vch2.Angle = Vch2.Angle + 360;

                    Vch3.Amp = V3rms;
                    Vch3.Angle = Phase3-Phase1;
                    if (Vch2.Angle < -360)
                        Vch3.Angle = Vch3.Angle + 360;

                    Ich4.Amp = Irms;
                    Ich4.Angle = Phase4-Phase1;
                    if (Ich4.Angle < -360)
                        Ich4.Angle = Ich4.Angle + 360;

                }                
            }             

            if (InvokeRequired)
            {
                Invoke((MethodInvoker)delegate
                {
                    /*以下是测量结果的显示*/
                    if (index==4)
                    {
                        index = 0;
                        textBox13.Text = V1rms.ToString("f2");
                        textBox1.Text = V2rms.ToString("f2");
                        textBox12.Text = V3rms.ToString("f2");
                        textBox20.Text = Vabrms.ToString("f2");
                        textBox17.Text = Vbcrms.ToString("f2");
                        textBox18.Text = Vcarms.ToString("f2");
                        textBox2.Text = Irms.ToString("f2");
                        textBox4.Text = P.ToString("f2");
                        textBox3.Text = S.ToString("f2");
                        textBox5.Text = Q.ToString("f2");
                        textBox6.Text = cos_fi.ToString("f2");
                        textBox7.Text = fi.ToString("f2");
                        textBox11.Text = S2.ToString("f2");
                        textBox10.Text = Q2.ToString("f2");
                        textBox9.Text = cos_fi2.ToString("f2");
                        textBox8.Text = fi2.ToString("f2");                       

                        easyChartX1.Plot(display, 0, 0.01); //显示波形
                        pictureBox1.Refresh();
                    }
                  
                });

                BeginInvoke((MethodInvoker)delegate
                {                 
                    
                    

                });
            }
        }
    }
}
