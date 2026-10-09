using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Web.Script.Serialization;

[assembly: AssemblyTitle("达人结算智能助手")]
[assembly: AssemblyDescription("达人结算核对、收款信息填报、异常检测及调差")]
[assembly: System.Runtime.Versioning.TargetFramework(".NETFramework,Version=v4.8")]
[assembly: AssemblyVersion("1.3.6.0")]
[assembly: AssemblyFileVersion("1.3.6.0")]

namespace CreatorSettlement {
public sealed class Cell { public string Text=""; public bool Formula,Error,Numeric; }
public sealed class Sheet {
 public string Name; public bool Date1904;
 public SortedDictionary<int,Dictionary<string,Cell>> Rows=new SortedDictionary<int,Dictionary<string,Cell>>();
 public int Header; public Dictionary<string,string> Fields=new Dictionary<string,string>();
 public string Get(int row,params string[] names) { foreach(string name in names) { string col; Cell c; if(Fields.TryGetValue(Xlsx.Norm(name),out col)&&Rows[row].TryGetValue(col,out c)) return c.Text.Trim(); } return ""; }
 public Cell GetCell(int row,params string[] names) { foreach(string name in names) {string col; Cell c;if(Fields.TryGetValue(Xlsx.Norm(name),out col)&&Rows[row].TryGetValue(col,out c))return c;}return new Cell(); }
 public bool Has(params string[] names) {return names.Any(n=>Fields.ContainsKey(Xlsx.Norm(n)));}
 public void SetHeader(int r) { Header=r;Fields.Clear();foreach(var c in Rows[r]) {string h=Xlsx.Norm(c.Value.Text);if(h.Length>0&&!Fields.ContainsKey(h))Fields[h]=c.Key;} }
}
public static class Xlsx {
 public static readonly XNamespace N="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
 static readonly CultureInfo CI=CultureInfo.InvariantCulture;
 public static string Norm(string s) {return Regex.Replace(s??"",@"\s+","").Replace("（","(").Replace("）",")");}
 public static string Col(int i){string s="";for(i++;i>0;i=(i-1)/26)s=(char)('A'+(i-1)%26)+s;return s;}
 static XDocument Doc(ZipArchive z,string path){var e=z.GetEntry(path);if(e==null)throw new Exception("Excel 缺少文件："+path);using(var s=e.Open())return XDocument.Load(s);}
 public static List<Sheet> Read(string path) {
  if(!path.EndsWith(".xlsx",StringComparison.OrdinalIgnoreCase))throw new Exception("请选择 .xlsx 文件。旧版 .xls 请先另存为 .xlsx。");
  using(var z=ZipFile.OpenRead(path)) {
   var shared=new List<string>();if(z.GetEntry("xl/sharedStrings.xml")!=null)shared=Doc(z,"xl/sharedStrings.xml").Root.Elements(N+"si").Select(x=>String.Concat(x.Descendants(N+"t").Select(t=>t.Value))).ToList();
   var wb=Doc(z,"xl/workbook.xml");var rel=Doc(z,"xl/_rels/workbook.xml.rels").Root.Elements().ToDictionary(x=>(string)x.Attribute("Id"),x=>(string)x.Attribute("Target"));
   bool d1904=wb.Root.Element(N+"workbookPr")!=null&&new[]{"1","true"}.Contains((string)wb.Root.Element(N+"workbookPr").Attribute("date1904"));
   var result=new List<Sheet>();foreach(var ws in wb.Root.Element(N+"sheets").Elements(N+"sheet")) {
    string target=rel[(string)ws.Attribute(XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships")+"id")].Replace('\\','/');
    string entry=target.StartsWith("/")?target.TrimStart('/'):"xl/"+target;entry=entry.Replace("xl/../","");
    var s=new Sheet{Name=(string)ws.Attribute("name"),Date1904=d1904};var xml=Doc(z,entry);
    foreach(var row in xml.Root.Element(N+"sheetData").Elements(N+"row")) {
     int r=(int)row.Attribute("r");var cells=new Dictionary<string,Cell>();
     foreach(var c in row.Elements(N+"c")) {string addr=(string)c.Attribute("r"),type=(string)c.Attribute("t"),v=(string)c.Element(N+"v")??"";string text=v;
      if(type=="s") {int k;if(!int.TryParse(v,out k)||k<0||k>=shared.Count)throw new Exception("Excel 共享字符串无效："+addr);text=shared[k];}
      if(type=="inlineStr")text=String.Concat(c.Descendants(N+"t").Select(t=>t.Value));
      cells[Regex.Replace(addr,@"\d","")]=new Cell{Text=text,Formula=c.Element(N+"f")!=null,Error=type=="e",Numeric=type==null||type=="n"};
     }s.Rows[r]=cells;
    }
    if(s.Rows.Count>0)s.SetHeader(s.Rows.First().Key);result.Add(s);
   }return result;
  }
 }
 static XElement MakeCell(string addr,string style,object value,string formula,decimal? cached) {
  var c=new XElement(N+"c",new XAttribute("r",addr));if(style!=null)c.SetAttributeValue("s",style);
  if(formula!=null){c.Add(new XElement(N+"f",formula));c.Add(new XElement(N+"v",cached.GetValueOrDefault().ToString(CI)));}
  else if(value!=null&&value!=DBNull.Value) {
   if(value is decimal||value is int||value is long||value is double)c.Add(new XElement(N+"v",Convert.ToString(value,CI)));
   else {string text=Convert.ToString(value,CI);if(text.Length>32767)throw new Exception("单元格文本超过 Excel 长度限制。");if(text.Length>0){c.SetAttributeValue("t","inlineStr");c.Add(new XElement(N+"is",new XElement(N+"t",new XAttribute(XNamespace.Xml+"space","preserve"),text)));}}
  }return c;
 }
 public static void Export(string template,string output,List<object[]> values,Func<int,int,string> formula,Func<int,int,decimal?> cache,bool footer) {
  if(values.Count>1048573)throw new Exception("输出超过 Excel 最大行数。");
  string tmp=output+"."+Guid.NewGuid().ToString("N")+".tmp";
  try {
   using(var src=Assembly.GetExecutingAssembly().GetManifestResourceStream(template+".xlsx"))using(var dst=File.Create(tmp)){if(src==null)throw new Exception("内置模板缺失。");src.CopyTo(dst);}
   using(var z=ZipFile.Open(tmp,ZipArchiveMode.Update)) {
    var doc=Doc(z,"xl/worksheets/sheet1.xml");var root=doc.Root;var sd=root.Element(N+"sheetData");var proto=sd.Elements(N+"row").First(x=>(int)x.Attribute("r")==2);var total=footer?new XElement(sd.Elements(N+"row").Last()):null;
    var styles=proto.Elements(N+"c").ToDictionary(c=>Regex.Replace((string)c.Attribute("r"),@"\d",""),c=>(string)c.Attribute("s"));var rowAttrs=proto.Attributes().Where(a=>a.Name!="r"&&a.Name!="spans").Select(a=>new XAttribute(a)).ToList();
    int templateCols=styles.Count;var oldDimension=root.Element(N+"dimension");if(oldDimension!=null){string lastAddress=((string)oldDimension.Attribute("ref")??"").Split(':').Last();int k=0;foreach(char ch in Regex.Replace(lastAddress,@"[^A-Z]",""))k=k*26+ch-'A'+1;templateCols=Math.Max(templateCols,k);}
    sd.Elements(N+"row").Where(r=>(int)r.Attribute("r")>1).Remove();
    for(int i=0;i<values.Count;i++) {int r=i+2;var row=new XElement(N+"row",new XAttribute("r",r),rowAttrs.Select(a=>new XAttribute(a)));int cols=Math.Max(templateCols,values[i].Length);
     for(int j=0;j<cols;j++){string col=Col(j),style;styles.TryGetValue(col,out style);row.Add(MakeCell(col+r,style,j<values[i].Length?values[i][j]:null,formula==null?null:formula(i,j),cache==null?null:cache(i,j)));}sd.Add(row);
    }
    int last=values.Count+1;
    if(footer){last++;total.SetAttributeValue("r",last);foreach(var c in total.Elements(N+"c")){string col=Regex.Replace((string)c.Attribute("r"),@"\d","");c.SetAttributeValue("r",col+last);if(col=="I"){c.RemoveNodes();c.Attribute("t").RemoveIfPresent();c.Add(new XElement(N+"f",values.Count==0?"0":"SUM(I2:I"+(last-1)+")"));decimal sum=0;for(int i=0;i<values.Count;i++)sum+=cache(i,8).GetValueOrDefault();c.Add(new XElement(N+"v",sum.ToString(CI)));}}sd.Add(total);}
    var dimension=root.Element(N+"dimension");if(dimension==null){dimension=new XElement(N+"dimension");root.AddFirst(dimension);}dimension.SetAttributeValue("ref","A1:"+Col(templateCols-1)+Math.Max(1,last));
var af=root.Element(N+"autoFilter");if(af==null){af=new XElement(N+"autoFilter");var following=root.Elements().FirstOrDefault(x=>new[]{"sortState","dataConsolidate","customSheetViews","mergeCells","phoneticPr","conditionalFormatting","dataValidations","hyperlinks","printOptions","pageMargins","pageSetup","headerFooter","rowBreaks","colBreaks","customProperties","cellWatches","ignoredErrors","smartTags","drawing","legacyDrawing","legacyDrawingHF","picture","oleObjects","controls","webPublishItems","tableParts","extLst"}.Contains(x.Name.LocalName));if(following==null)root.Add(af);else following.AddBeforeSelf(af);}int filterCols=template=="check"?10:template=="payment"?13:8;af.SetAttributeValue("ref","A1:"+Col(filterCols-1)+Math.Max(1,values.Count+1));
    foreach(var view in root.Descendants(N+"sheetView")){view.SetAttributeValue("topLeftCell","A1");foreach(var sel in view.Elements(N+"selection")){sel.SetAttributeValue("activeCell","A1");sel.SetAttributeValue("sqref","A1");}}
    Replace(z,"xl/worksheets/sheet1.xml",doc);
    var wb=Doc(z,"xl/workbook.xml");wb.Root.Elements(N+"calcPr").Remove();wb.Root.Add(new XElement(N+"calcPr",new XAttribute("calcMode","auto"),new XAttribute("fullCalcOnLoad","1")));Replace(z,"xl/workbook.xml",wb);
   }
   // Caller obtains overwrite consent; final write occurs only after the package is complete.
   if(File.Exists(output))File.Replace(tmp,output,null,true);else File.Move(tmp,output);
  }finally{if(File.Exists(tmp))File.Delete(tmp);}
 }
 static void Replace(ZipArchive z,string path,XDocument d){z.GetEntry(path).Delete();using(var s=z.CreateEntry(path).Open())d.Save(s);}
 public static void RemoveIfPresent(this XAttribute a){if(a!=null)a.Remove();}
}
public sealed class Record {
 public int Row;public string Sheet,RecordId,Task,Name,Type,Note;
 public DateTime Date;public decimal Count,Price,Paid;
 public decimal Cost {get{return Money(Count*Price);}}public static decimal Money(decimal n){return Decimal.Round(n,2,MidpointRounding.AwayFromZero);}
}
public sealed class Issue {public string Category,Task,Source,Detail;}
public sealed class SavedSession {public int Version=2;public List<Dictionary<string,string>> Payments;public List<Dictionary<string,string>> Overrides;}
public sealed class Receipt {public string Channel="",Account="",Identity="",Name="",Phone="",Target=null;}
public sealed class Segment {
 public string Task,Name,Type,Note;public DateTime Start,End;public decimal Price,Count;
 public decimal Cost {get{return Record.Money(Count*Price);}}
 public string Range {get {return Start==End?Start.ToString("M.d"):Start.ToString("M.d")+"-"+End.ToString("M.d");}}
}
public sealed partial class Model {
 public List<Record> Records=new List<Record>();public List<Issue> Issues=new List<Issue>();public List<Segment> Segments=new List<Segment>();
 public DataTable Check,Payment,Adjust,Warnings;
 static decimal Number(Cell c,string label){decimal n;if(c.Error||c.Formula)throw new Exception(label+" 含 Excel 错误或公式。请在副本中粘贴为数值后导入。");if(!Decimal.TryParse(c.Text,NumberStyles.Number|NumberStyles.AllowExponent,CultureInfo.InvariantCulture,out n))throw new Exception(label+" 为空或不是有效数字。");return n;}
 void Warn(string category,string task,string source,string detail){Issues.Add(new Issue{Category=category,Task=task,Source=source,Detail=detail});}
 public static Model Load(Sheet s) {
  foreach(string f in new[]{"任务ID","结算日期","任务名称","任务类型","转化数","单价","打款金额"})if(!s.Has(f))throw new Exception("工作表缺少列："+f+"。请选择原始结算明细工作表。");
  var m=new Model();foreach(int row in s.Rows.Keys.Where(r=>r>s.Header)){
   if(s.Rows[row].Values.All(c=>String.IsNullOrWhiteSpace(c.Text)))continue;
   string task=s.Get(row,"任务ID");if(String.IsNullOrEmpty(task))throw new Exception(s.Name+" 第 "+row+" 行任务ID为空。");
   DateTime date;var dc=s.GetCell(row,"结算日期");string dt=dc.Text;
   if(dc.Formula||dc.Error)throw new Exception("第 "+row+" 行日期含公式或错误。");
   if(!DateTime.TryParse(dt,CultureInfo.InvariantCulture,DateTimeStyles.None,out date)){double serial;if(!double.TryParse(dt,NumberStyles.Float,CultureInfo.InvariantCulture,out serial)||serial<1||serial>2958465)throw new Exception("第 "+row+" 行结算日期无效："+dt);date=DateTime.FromOADate(serial+(s.Date1904?1462:0));}
   var r=new Record{Row=row,Sheet=s.Name,RecordId=s.Get(row,"ID"),Task=task,Name=s.Get(row,"任务名称"),Type=s.Get(row,"任务类型"),Date=date.Date,Count=Number(s.GetCell(row,"转化数"),"第 "+row+" 行转化数"),Price=Number(s.GetCell(row,"单价"),"第 "+row+" 行单价"),Paid=Number(s.GetCell(row,"打款金额"),"第 "+row+" 行打款金额"),Note=s.Get(row,"备注")};m.Records.Add(r);
   if(r.Count<0||r.Count!=Decimal.Truncate(r.Count)||r.Price<0||r.Paid<0)m.Warn("金额异常",task,s.Name+":"+row,"转化数应为非负整数，单价及金额应为非负数。");
   if(r.Cost!=r.Paid)m.Warn("金额异常",task,s.Name+":"+row,"应计 "+r.Cost+"，原始打款 "+r.Paid+"，差额 "+(r.Cost-r.Paid));
   var idCell=s.GetCell(row,"任务ID");if(idCell.Numeric&&idCell.Text.Replace(".","").Length>15)m.Warn("标识精度",task,s.Name+":"+row,"任务ID为长数字，请将原始列设为文本后核实精度。");
  }
  if(m.Records.Count==0)throw new Exception("工作表没有结算记录。");
  foreach(var duplicate in m.Records.Where(r=>r.RecordId!="").GroupBy(r=>r.RecordId).Where(g=>g.Count()>1))m.Warn("重复任务",duplicate.First().Task,s.Name+":"+String.Join(",",duplicate.Select(r=>r.Row)),"原始记录ID重复："+duplicate.Key+"。保留全部记录，未自动去重。");
  foreach(var g in m.Records.GroupBy(r=>r.Task)){
   if(g.Select(r=>r.Price).Distinct().Count()>1)m.Warn("同任务ID不同价格",g.Key,s.Name,"单价："+String.Join("、",g.Select(r=>r.Price).Distinct().OrderBy(x=>x)));
   var dates=g.Select(r=>r.Date).Distinct().OrderBy(x=>x).ToList();for(int i=1;i<dates.Count;i++)if((dates[i]-dates[i-1]).Days>1)m.Warn("日期断档",g.Key,s.Name,dates[i-1].AddDays(1).ToString("yyyy-MM-dd")+" 至 "+dates[i].AddDays(-1).ToString("yyyy-MM-dd")+" 无记录；同单价日期段跨断档合并，转化数仅汇总实际记录。");
   foreach(var dup in g.GroupBy(r=>new{r.Date,r.Price}).Where(x=>x.Count()>1))m.Warn("重复任务",g.Key,s.Name+":"+String.Join(",",dup.Select(r=>r.Row)),"同任务、同日期、同单价有多行。保留并汇总，需人工核实。");
   if(g.Select(r=>r.Name+"|"+r.Type).Distinct().Count()>1)m.Warn("任务信息冲突",g.Key,s.Name,"同任务关联多个名称或类型，请核实。");
   var daily=g.GroupBy(r=>new{r.Date,r.Price}).OrderBy(x=>x.Key.Date).ThenBy(x=>x.Key.Price);
   Segment current=null;foreach(var d in daily){var first=d.First();if(current==null||current.Price!=d.Key.Price){current=new Segment{Task=g.Key,Name="",Type="",Start=d.Key.Date,End=d.Key.Date,Price=d.Key.Price,Note=""};m.Segments.Add(current);}current.End=d.Key.Date;current.Count+=d.Sum(r=>r.Count);current.Name=JoinDistinct(new[]{current.Name}.Concat(d.Select(r=>r.Name)));current.Type=JoinDistinct(new[]{current.Type}.Concat(d.Select(r=>r.Type)));current.Note=String.Join("；",new[]{current.Note}.Concat(d.Select(r=>r.Note)).Where(x=>x!="").Distinct());}
  }
  m.MakeTables();return m;
 }
 static DataTable Table(string[] names,params string[] numeric){var t=new DataTable();foreach(string n in names)t.Columns.Add(n,numeric.Contains(n)?typeof(decimal):typeof(string));return t;}
 void MakeTables(){
  Check=Table(new[]{"结算时间","任务/方案ID","方案名","任务类型","转化数","单价","打款金额","一口价（无手续费）","实际打款","备注"},"转化数","单价","打款金额","一口价（无手续费）","实际打款");
  foreach(var s in Segments)Check.Rows.Add(s.Range,s.Task,s.Name,s.Type,s.Count,s.Price,s.Cost,DBNull.Value,s.Cost,s.Note);
  Payment=Table(new[]{"任务ID","达人ID","任务名称","类型(必填)","收款平台(必填)","收款类型(必填)","打款对象类型(必填)","团长比例","团长打款金额","银行卡号","身份证号(必填)","姓名(必填)","支付宝账号","手机号(必填)"});
  Payment.Columns.Add("收款账号",typeof(string),"IIF([收款平台(必填)] = '银行卡', [银行卡号], [支付宝账号])");
  foreach(var g in Records.GroupBy(r=>r.Task)){
   var p=Payment.NewRow();p["任务ID"]=g.Key;p["任务名称"]=g.First().Name;p["类型(必填)"]=g.First().Type;p["收款类型(必填)"]="小额";p["打款对象类型(必填)"]="团长";p["团长比例"]="1";Payment.Rows.Add(p);
  }
  Adjust=Table(new[]{"任务类型","聚星任务名","优投方案ID","任务ID","线上成本","实际扣款金额","下单账户ID","备注"},"线上成本");
  foreach(var group in Records.Where(r=>r.Count*r.Price!=r.Paid).GroupBy(r=>r.Task)) {
   string types=String.Join("；",group.Select(r=>r.Type).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct());
   string names=String.Join("；",group.Select(r=>r.Name).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct());
   string notes=String.Join("；",group.Select(r=>r.Note).Where(x=>!String.IsNullOrWhiteSpace(x)).Distinct());
   Adjust.Rows.Add(types,names,"",group.Key,group.Sum(r=>r.Count*r.Price-r.Paid),DBNull.Value,"",notes);
  }
  Warnings=Table(new[]{"异常类型","任务ID","来源工作表/行","说明"});foreach(var i in Issues)Warnings.Rows.Add(i.Category,i.Task,i.Source,i.Detail);
  Refresh();
 }
 public void Refresh(){
  foreach(DataRow r in Check.Rows){decimal fixedPrice=r.IsNull(7)?0:Convert.ToDecimal(r[7]);decimal cost=Record.Money((r.IsNull(4)?0:Convert.ToDecimal(r[4]))*(r.IsNull(5)?0:Convert.ToDecimal(r[5])));r[6]=cost;if(fixedPrice<0)throw new Exception("任务 "+r[1]+" 的一口价不能小于 0。");r[8]=Record.Money(cost-fixedPrice);}
 }
 public static string Val(DataRow r,string c){return r.IsNull(c)?"":Convert.ToString(r[c]).Trim();}
 public string Period {get{return Records.Min(r=>r.Date).ToString("yyyy.M.d")+"-"+Records.Max(r=>r.Date).ToString("yyyy.M.d");}}
 public string CheckFileName {get{return "结算核对_"+Period;}}
 public string ReportName(int index){return new[]{"结算核对","达人线下成本-打款信息填报","异常报告","结算调差表"}[index]+"_"+Period;}
 public void ApplyReceipt(Receipt info,IEnumerable<string> tasks=null){
  var ids=tasks==null?null:new HashSet<string>(tasks);foreach(DataRow p in Payment.Rows){if(ids!=null&&!ids.Contains(Val(p,"任务ID")))continue;
   p["收款平台(必填)"]=info.Channel.Trim();p["银行卡号"]=info.Channel=="银行卡"?info.Account.Trim():"";p["支付宝账号"]=info.Channel=="支付宝"?info.Account.Trim():"";
   p["身份证号(必填)"]=info.Identity.Trim();p["姓名(必填)"]=info.Name.Trim();p["手机号(必填)"]=info.Phone.Trim();
   p["收款类型(必填)"]="小额";p["团长比例"]="1";
   if(info.Target!=null)p["打款对象类型(必填)"]=info.Target.Trim();
   foreach(string field in new[]{"达人ID","团长打款金额"})p[field]=DBNull.Value;
  }
 }
 public void ValidatePayment(){var errors=new List<string>();foreach(DataRow r in Payment.Rows){string id=Val(r,"任务ID"),channel=Val(r,"收款平台(必填)"),bank=Val(r,"银行卡号"),alipay=Val(r,"支付宝账号");
   if(!new[]{"团长","达人"}.Contains(Val(r,"打款对象类型(必填)")))errors.Add(id+"：打款对象请选择团长或达人");
   if(channel!=""&&channel!="银行卡"&&channel!="支付宝")errors.Add(id+"：收款平台请选择银行卡或支付宝");
   if(bank!=""&&channel=="银行卡"&&!Regex.IsMatch(bank,@"^\d{10,30}$"))errors.Add(id+"：银行卡号应为10至30位数字");
   if(channel==""&&(bank!=""||alipay!=""))errors.Add(id+"：填写账号时请先选择收款平台");
   if(Val(r,"身份证号(必填)")!=""&&!Regex.IsMatch(Val(r,"身份证号(必填)"),@"^(\d{15}|\d{17}[0-9Xx])$"))errors.Add(id+"：身份证号码格式无效");
   if(Val(r,"手机号(必填)")!=""&&!Regex.IsMatch(Val(r,"手机号(必填)"),@"^1\d{10}$"))errors.Add(id+"：手机号应为11位数字");
  }if(errors.Count>0)throw new Exception("请检查已填写的信息：\n"+String.Join("\n",errors.Take(8)));
 }
 public int ImportInfo(Sheet s){
  if(!s.Has("任务ID"))throw new Exception("批量导入文件需要任务ID列。统一收款人可直接在左侧面板填写。");
  var aliases=new Dictionary<string,string[]>{{"收款平台(必填)",new[]{"收款平台(必填)","收款方式","收款平台"}},{"银行卡号",new[]{"银行卡号","银行卡"}},{"身份证号(必填)",new[]{"身份证号(必填)","身份证号","身份证"}},{"姓名(必填)",new[]{"姓名(必填)","姓名","收款方姓名"}},{"支付宝账号",new[]{"支付宝账号","支付宝"}},{"手机号(必填)",new[]{"手机号(必填)","手机号","手机"}}};
  var rows=s.Rows.Keys.Where(r=>r>s.Header&&!s.Rows[r].Values.All(c=>c.Text.Trim()=="")).ToList();var plan=new List<Tuple<DataRow,Receipt>>();
  foreach(DataRow p in Payment.Rows){string task=Val(p,"任务ID");var matches=rows.Where(r=>s.Get(r,"任务ID")==task).ToList();if(matches.Count==0)continue;
   var signatures=matches.Select(r=>String.Join("|",aliases.Values.Select(a=>s.Get(r,a)))).Distinct();if(signatures.Count()>1)throw new Exception("任务 "+task+" 的收款信息有冲突，来源行："+String.Join(",",matches));
   int row=matches[0];foreach(var a in aliases){var cell=s.GetCell(row,a.Value);if(cell.Formula||cell.Error)throw new Exception("第 "+row+" 行 "+a.Key+" 含公式或错误，请粘贴为文本。");if(a.Key!="姓名(必填)"&&cell.Numeric&&cell.Text.Replace(".","").Length>15)throw new Exception("第 "+row+" 行 "+a.Key+" 是长数字，请核实精度并以文本格式填写。");}
   string channel=s.Get(row,aliases["收款平台(必填)"]);plan.Add(Tuple.Create(p,new Receipt{Channel=channel,Account=s.Get(row,aliases[channel=="支付宝"?"支付宝账号":"银行卡号"]),Identity=s.Get(row,aliases["身份证号(必填)"]),Name=s.Get(row,aliases["姓名(必填)"]),Phone=s.Get(row,aliases["手机号(必填)"])}));
  }
  if(plan.Count==0)throw new Exception("没有匹配到任务ID，请检查导入文件。");foreach(var item in plan)ApplyReceipt(item.Item2,new[]{Val(item.Item1,"任务ID")});return plan.Count;
 }
 public void ExportCheck(string path){Refresh();var rows=Check.AsEnumerable().Select(r=>r.ItemArray).ToList();Xlsx.Export("check",path,rows,(i,j)=>j==6?"E"+(i+2)+"*F"+(i+2):j==8?"G"+(i+2)+"-H"+(i+2):null,(i,j)=>j==6?(decimal?)Convert.ToDecimal(Check.Rows[i][6]):j==8?(decimal?)Convert.ToDecimal(Check.Rows[i][8]):null,true);}
 public void ExportPayment(string path){ValidatePayment();var rows=new List<object[]>();foreach(DataRow r in Payment.Rows){string channel=Val(r,"收款平台(必填)");
   rows.Add(new object[]{r["类型(必填)"],r["任务ID"],null,channel,"小额",r["打款对象类型(必填)"],r.IsNull("团长比例")||Val(r,"团长比例")==""?(object)null:Decimal.Parse(Val(r,"团长比例"),CultureInfo.InvariantCulture),r.IsNull("团长打款金额")||Val(r,"团长打款金额")==""?(object)null:Decimal.Parse(Val(r,"团长打款金额"),CultureInfo.InvariantCulture),channel=="银行卡"?r["银行卡号"]:null,r["身份证号(必填)"],r["姓名(必填)"],channel=="支付宝"?r["支付宝账号"]:null,r["手机号(必填)"]});}
  Xlsx.Export("payment",path,rows,null,null,false);
 }
 public void ExportAdjust(string path){Xlsx.Export("adjust",path,Adjust.AsEnumerable().Select(r=>r.ItemArray).ToList(),null,null,false);}
 public void ExportWarnings(string path){
  // Reuse the check template's styles but replace its header, data and footer in a separate package.
  var rows=Warnings.AsEnumerable().Select(r=>r.ItemArray).ToList();Xlsx.Export("adjust",path,rows,null,null,false);
  using(var z=ZipFile.Open(path,ZipArchiveMode.Update)){var e=z.GetEntry("xl/worksheets/sheet1.xml");XDocument d;using(var stream=e.Open())d=XDocument.Load(stream);var sd=d.Root.Element(Xlsx.N+"sheetData");var header=sd.Elements(Xlsx.N+"row").First();header.RemoveNodes();for(int j=0;j<4;j++)header.Add(new XElement(Xlsx.N+"c",new XAttribute("r",Xlsx.Col(j)+"1"),new XAttribute("s","6"),new XAttribute("t","inlineStr"),new XElement(Xlsx.N+"is",new XElement(Xlsx.N+"t",Warnings.Columns[j].ColumnName))));var cols=d.Root.Element(Xlsx.N+"cols");cols.RemoveNodes();double[] widths={24,27,28,90};for(int j=0;j<4;j++)cols.Add(new XElement(Xlsx.N+"col",new XAttribute("min",j+1),new XAttribute("max",j+1),new XAttribute("width",widths[j]),new XAttribute("customWidth","1")));foreach(var row in sd.Elements(Xlsx.N+"row").Skip(1)){row.SetAttributeValue("ht","65");row.SetAttributeValue("customHeight","1");row.Elements(Xlsx.N+"c").Skip(4).Remove();}d.Root.Element(Xlsx.N+"dimension").SetAttributeValue("ref","A1:D"+(Warnings.Rows.Count+1));e.Delete();using(var stream=z.CreateEntry("xl/worksheets/sheet1.xml").Open())d.Save(stream);}
 }
}

public static class Program {
 [STAThread] public static int Main(string[] args){try{if(args.Length>=3&&args[0]=="--verify"){var s=Xlsx.Read(args[1]).First();foreach(var row in s.Rows.Take(30))if(row.Value.Values.Any(c=>Xlsx.Norm(c.Text)=="任务ID")){s.SetHeader(row.Key);break;}var m=Model.Load(s);Directory.CreateDirectory(args[2]);m.ExportCheck(Path.Combine(args[2],"check.xlsx"));m.ExportAdjust(Path.Combine(args[2],"adjust.xlsx"));m.ExportWarnings(Path.Combine(args[2],"warnings.xlsx"));File.WriteAllText(Path.Combine(args[2],"stats.txt"),"records="+m.Records.Count+"\ntasks="+m.Payment.Rows.Count+"\nsegments="+m.Check.Rows.Count+"\ncount="+m.Records.Sum(r=>r.Count)+"\ncost="+m.Records.Sum(r=>r.Cost)+"\npaid="+m.Records.Sum(r=>r.Paid)+"\nadjust="+m.Adjust.Rows.Count,Encoding.UTF8);return 0;}
  Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);Application.ThreadException+=(o,e)=>MessageBox.Show(e.Exception.Message,"操作失败");Application.Run(new MainWindow());return 0;
 }catch(Exception e){if(args.Length>0)File.WriteAllText(Path.Combine(Path.GetTempPath(),"creator-settlement-verify-error.txt"),e.ToString());else MessageBox.Show(e.Message,"启动失败");return 1;}}
}
}
