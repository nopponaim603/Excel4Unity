using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;

public class Example : MonoBehaviour {
    public Text outputText;
	// Use this for initialization
	void Start () {
        string excelInPutFileName = "Test.xlsx";
        string excelOutPutFileName = "Test2.xlsx";
        string excelInPutFilePath = System.IO.Path.Combine(Application.streamingAssetsPath, excelInPutFileName);
        string excelOutPutFilePath = System.IO.Path.Combine(Application.streamingAssetsPath, excelOutPutFileName);
        
        Excel xls = ExcelHelper.LoadExcel(excelInPutFilePath);
        xls.ShowLog();
        print(xls.Tables.Count);

        xls.Tables[0].SetValue(1, 1, "???");
        ExcelHelper.SaveExcel(xls, excelOutPutFilePath);
        outputText.text = "Save Success.";
    }

}
