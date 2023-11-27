namespace Parce
{
    class ElevatorModel
    {
        public string _fOrgCode { get; set; } = "33";
        public string _sOrgCode { get; set; }
        public string _api { get; } = "3401000000";
        public string _street { get; set; }
        public string _home { get; set; }
        public string _date { get; set; }
        public string _serviceCode { get; } = "11109";
        public string _serviceName { get; } = "Техническое обслуживание";
        public string _counterId { get; } = "0";
        public string _firstCounterValue { get; } = "0";
        public string _lastCounterValue { get; } = "0";
        public string _serviceScope { get; } = "0";
        public string _ratio { get; } = "0";
        public string _rate { get; } = "0";
        public string _summ { get; set; }
        public string _rateACT { get; } = "0";
        public string _ACTsumm { get; } = "0";
        public string _summWhithACT { get; set; }
    }
}
