using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMSync.Classes
{
    class EnumSync
    {
        public enum YesNo
        {
            Yes = 1,
            No = 0
        }

        public enum Status
        {
            Active = 1,
            Inactive = 0
        }

        public enum PioneerStatus
        {
            Pioneer = 1,
            ITA = 0
        }

        public enum Gender
        {
            Male = 0,
            Female = 1
        }

        public enum RelocationStatus
        {
            Null = 0,
            Delocated = 1,
            Exemption = 2,
            NotRelocated = 3,
            Relocated = 4,
            Revoked = 5,
            ToBeRevoked = 6,
            Token = 7,
            Under6MonthsGracePeriod = 8,
            UnderExtension = 9
        }

        public enum AccessMode
        {
            Public = 1,
            Private = 0
        }

        public enum OperationalStatus
        {
            Null = 0,
            Closed = 1,
            Uncontactable = 2,
            Surrendering = 3,
            RevokedButUnofficial = 4,
            Revoked = 5,
            Active = 6,
            Merged = 7,
            Dormant = 8,
            Unincorporated = 9,
            Others = 10,
            Surrendered = 11
        }
        public enum TraceLogType
        {
            Workflow = 3001,
            AccountCategory = 3002,
            OperationalStatus = 3003,
            CustomerRanking = 3004
        }
    }
}
