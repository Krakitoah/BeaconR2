using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;

namespace BeaconProtoType
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private void EntryCell_Completed(object sender, EventArgs e)
        {
            if (Name.Text.ToLower() == "Test".ToLower())
                Navigation.PushAsync(new TimeStampPage());
            else
            {
                DisplayAlert("Login Incomplete", "That is not a valid user", "Ok");
            }
        }
    }
}
