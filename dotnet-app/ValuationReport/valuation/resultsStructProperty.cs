public class resultsStructProperty
{
    public string fileName;
    public string section;
    public string description;
    public string issueDescription;
    public string deviation;
  }

public class monthlyStruct
{
    public string client;
    public cashproperty cashproperty;
    public valuationproperty valuationproperty;
    public acquisitionDisposals acquisitionDisposals;
    public valuationSummaryproperty valuationSummary;
}

public class quarterlyStruct
{
    public string client;
    public cashproperty cashproperty;
    public valuationproperty valuationproperty;
    public acquisitionDisposals acquisitionDisposals;
    public performanceproperty performance;
    public invoiceproperty invoice;
    public reconcileValuesProperty reconcileValues;
}

public class exceptionStruct
{
    public string fileName;
    public string problemArea;
    public string description;
}

public class bidClientException
{
    public bool isBidClient;
    public string clientName;
}

public class reconcileValuesProperty
{
    public string sectionName;
    public double totalMarketValue;
    public bool valuesMatchCheck;
}