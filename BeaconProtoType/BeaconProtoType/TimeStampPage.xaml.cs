using Android.Bluetooth;
using Android.Bluetooth.LE;
using Android.Content;
using Android.Runtime;
using BeaconProtoType;
using Plugin.BluetoothLE;
using Plugin.Messaging;
using SQLite;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace BeaconProtoType
{
    [XamlCompilation(XamlCompilationOptions.Compile)]
    public partial class TimeStampPage : ContentPage
    {

        public ObservableCollection<TimeInPunches> myTimeIns;
        public ObservableCollection<TimeInPunches> myTimeOuts;
        public ObservableCollection<IDevice> myDevices;
        public ObservableCollection<String> deviceNames;
        private SQLiteAsyncConnection connection;
        private BluetoothAdapter mBluetoothAdapter;
        public float rssi;


        public TimeStampPage()
        {
            InitializeComponent();

            connection = DependencyService.Get<ISQLiteDb>().GetConnection();
            mBluetoothAdapter = BluetoothAdapter.DefaultAdapter;

        }

        protected override async void OnAppearing()
        {
            await connection.CreateTableAsync<TimeInPunches>();

            var myTimes = await connection.Table<TimeInPunches>().ToListAsync();

            myTimeIns = new ObservableCollection<TimeInPunches>(myTimes);
            myTimeOuts = new ObservableCollection<TimeInPunches>(myTimes);
            myDevices = new ObservableCollection<IDevice>();
            deviceNames = new ObservableCollection<string>();

            TimeInList.ItemsSource = myTimeIns;

            DeviceList.ItemsSource = deviceNames;

            if (!mBluetoothAdapter.IsEnabled)
            {
                var result = await DisplayAlert("Bluetooth Acivation", "This device will require bluetooth activation, may I activate bluetooth?", "Ok", "No");

                if (result)
                {
                    mBluetoothAdapter.Enable();
                }

            }

            base.OnAppearing();

        }

        private async void TimeIn_Clicked(object sender, EventArgs e)
        {
            var Punch = new TimeInPunches
            {
                UserName = "test",
                BeaconID = "test",
                PhoneID = "test",
                TimeIn = DateTime.Now,
                SignalStrength = "close"
            };
            if (myTimeIns.Count == myTimeOuts.Count)
            {
                await connection.InsertAsync(Punch);
                myTimeIns.Add(Punch);
            }
            else
            {
                await DisplayAlert("Invalid TimePunch", "Cannot Punch In Without Punching Out", "Ok");
            }
        }

        private async void TimeOut_Clicked(object sender, EventArgs e)
        {
            var upDatedPunch = myTimeIns[0];
            upDatedPunch.TimeOut = DateTime.Now;

            if (myTimeOuts.Count == myTimeIns.Count - 1)
            {
                await connection.UpdateAsync(upDatedPunch);
            }

            else
            {
                await DisplayAlert("Invalid TimePunch", "Cannot Punch Out Without Punching In", "Ok");
            }

        }

        private async void Convert_Clicked(object sender, EventArgs e)
        {

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2013;

                IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);

                IWorksheet worksheet = workbook.Worksheets[0];


                //worksheet["A1"].Text = "User";
                //worksheet["B1"].Text = "PhoneID";
                //worksheet["C1"].Text = "TimeIn";
                //worksheet["D1"].Text = "TimeOut";
                //worksheet["E1"].Text = "BeaconID";
                //worksheet["F1"].Text = "SignalStrength";


                for (int i = 0; i < deviceNames.Count; i++)
                {
                    //worksheet["A" + Convert.ToString(i + 2)].Text = myTimeIns[i].UserName;
                    //worksheet["B" + Convert.ToString(i + 2)].Text = myTimeIns[i].PhoneID;
                    //worksheet["C" + Convert.ToString(i + 2)].Text = myTimeIns[i].TimeIn.ToString();
                    //worksheet["D" + Convert.ToString(i + 2)].Text = myTimeIns[i].TimeOut.ToString();
                    //worksheet["E" + Convert.ToString(i + 2)].Text = myTimeIns[i].BeaconID;
                    //worksheet["F" + Convert.ToString(i + 2)].Text = myTimeIns[i].SignalStrength;
                    worksheet["A" + Convert.ToString(i + 3)].Text = deviceNames[i];


                }



                worksheet["A1:F1"].ColumnWidth = 10;

                IRange headingRange = worksheet["A1:E1"];
                headingRange.CellStyle.Font.Bold = true;
                headingRange.CellStyle.ColorIndex = ExcelKnownColors.Light_green;

                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                workbook.Close();

                await Xamarin.Forms.DependencyService.Get<ISave>().SaveAndView("GettingStared.xlsx", "application/msexcel", stream);
            }
        }

        private async void ClearDatabase_Clicked(object sender, EventArgs e)
        {
            await connection.DropTableAsync<TimeInPunches>();
            await connection.CreateTableAsync<TimeInPunches>();
            myTimeIns.Clear();
            myTimeOuts.Clear();
        }

        private async void TimeInList_ItemSelected(object sender, SelectedItemChangedEventArgs e)
        {
            var mySelection = e.SelectedItem as TimeInPunches;
            await DisplayAlert("Alert", Convert.ToString(mySelection.TimeOut), "ok");
        }

        private void ScanForBluetooth_Clicked(object sender, EventArgs e)
        {
            CrossBleAdapter.Current.Scan().Subscribe(scanResult =>
            {
                if (!string.IsNullOrWhiteSpace(scanResult.Device.Name) && !scanResult.Device.Name.Contains("Alta") && scanResult.Rssi >= -60)
                {
                    deviceNames.Add(String.Format("{0} - RSSI: {1} Time: {2}", scanResult.Device.Name, scanResult.Rssi, DateTime.Now));
                }
            });
        }

        private async void TimeOutList_ItemSelected(object sender, SelectedItemChangedEventArgs e)
        {
            var mySelection = e.SelectedItem as string;
            await DisplayAlert(mySelection, mySelection, "ok");
        }
    }
}


