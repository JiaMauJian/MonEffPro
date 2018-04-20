# MonEffPro

# 透過LINQPad學習LINQ
* 安裝LINQPad
* 安裝LINQPad-> IQDriver (連Oracle使用)
![alt text](https://github.com/adam-p/markdown-here/raw/master/src/common/images/icon48.png "Logo Title Text 1")
![alt text](https://github.com/adam-p/markdown-here/raw/master/src/common/images/icon48.png "Logo Title Text 1")
* 擴充Table -> Class的Extension Function (LINQPadExtensions.cs 把Code貼到LINQPad My Extensions裡面，然後按F5)
* 按F4加入Dapper.dll參考以及增加Additional Namespace Imports (Dapper.dll可以用VS2017 Nuget安裝後再去把它摳出來)
```
Dapper
System.Data.Entity
```
* SQL產生Class
```
this.Connection.DumpClass("SELECT * FROM EDWADM.MEDA_MMS_ipa_P@DBLINK_EDWUSER_28").Dump();
```
![alt text](https://github.com/adam-p/markdown-here/raw/master/src/common/images/icon48.png "Logo Title Text 1")

* SQL產生LINQ Object
```
void Main()
{	
	using (var connection = this.Connection)
	{			
		var sqlCmd = "SELECT * FROM EDWADM.MEDA_MMS_ipa_P@DBLINK_EDWUSER_28 WHERE ROWNUM <= 100";
		var result = connection.Query<MEDA_MMS_ipa_P>(sqlCmd);
		
		var x = from m in result where m.IPA_MACHID_LIST == "IPA-02" select m;
		x.Dump();
	}
}

public class MEDA_MMS_ipa_P
{
  public decimal? IPA_MONTHKEY { get; set; }

  public decimal? IPA_WEEKKEY { get; set; }

  public string IPA_DAYKEY { get; set; }

  public string IPA_FAB { get; set; }

  public string LINE { get; set; }

  public string WO_ID { get; set; }

  public string LOT_ID { get; set; }

  public DateTime? IPA_BOOKINTIME { get; set; }

  public DateTime? IPA_BOOKOUTTIME { get; set; }

  public string IPA_MACHID_LIST { get; set; }

  public int? IPA_預計投入 { get; set; }

  public int? IPA_實際投入 { get; set; }

  public int? IPA_產出數 { get; set; }

  public int? IPA_工單差 { get; set; }

  public int? IPA_總破片 { get; set; }

  public int? IPA_總缺角 { get; set; }

  public int? IPA_總退料 { get; set; }

  public int? IPA_總重工_IR1 { get; set; }

  public int? IPA_總重工_IR2 { get; set; }

  public int? IPA_待處理數量 { get; set; }

  public int? IPA_機台投入數 { get; set; }

  public int? IPA_未投破 { get; set; }

  public int? IPA_已投破 { get; set; }

  public int? IPA_未投缺 { get; set; }

  public int? IPA_已投缺 { get; set; }

  public int? IPA_未投重工_IR1 { get; set; }

  public int? IPA_已投重工_IR1 { get; set; }

  public int? IPA_未投重工_IR2 { get; set; }

  public int? IPA_已投重工_IR2 { get; set; }

  public int? IPA_未投退 { get; set; }

  public int? IPA_已投退 { get; set; }

  public int? IPA_不良品 { get; set; }

  public string IPA_SHIFT { get; set; }

  public string IPA_EMPID { get; set; }

  public string IPA_EMPNAME { get; set; }

  public string IPA_REMARK_LIST { get; set; }

  public DateTime? TIMESTAMP { get; set; }

  public string IPA_UPDATETIME { get; set; }

}
```
![alt text](https://github.com/adam-p/markdown-here/raw/master/src/common/images/icon48.png "Logo Title Text 1")
