using SQLite;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BeaconProtoType
{
    public class TimeInPunches
    {
        public string UserName { get; set; }
        [PrimaryKey, AutoIncrement]
        public int Id { get; set; }
        public string PhoneID { get; set; }
        public DateTime TimeIn { get; set; }

        public DateTime TimeOut { get; set; }
        public string BeaconID { get; set; }
        public string SignalStrength { get; set; }

    }
}
