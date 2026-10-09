using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CreatorSettlement {
public sealed partial class Model {
 public static string JoinDistinct(IEnumerable<string> values){return String.Join("；",values.Where(x=>!String.IsNullOrWhiteSpace(x)).SelectMany(x=>x.Split('；')).Distinct());}
 public int SyncFixedToAdjustment(){
  Refresh();var fixedGroups=Check.AsEnumerable().Where(r=>!r.IsNull(7)).GroupBy(r=>Val(r,"任务/方案ID")).ToList();
  var replacement=new HashSet<string>(fixedGroups.Select(g=>g.Key));var rows=new List<object[]>();
  foreach(var g in Records.Where(r=>r.Count*r.Price!=r.Paid).GroupBy(r=>r.Task))if(!replacement.Contains(g.Key))rows.Add(new object[]{JoinDistinct(g.Select(r=>r.Type)),JoinDistinct(g.Select(r=>r.Name)),"",g.Key,g.Sum(r=>r.Count*r.Price-r.Paid),DBNull.Value,"",JoinDistinct(g.Select(r=>r.Note))});
  foreach(var g in fixedGroups)rows.Add(new object[]{JoinDistinct(g.Select(r=>Val(r,"任务类型"))),JoinDistinct(g.Select(r=>Val(r,"方案名"))),"",g.Key,g.Sum(r=>(decimal)r[7]),DBNull.Value,"",JoinDistinct(g.Select(r=>Val(r,"备注")))});
  Adjust.BeginLoadData();try{Adjust.Clear();foreach(var r in rows)Adjust.Rows.Add(r);}finally{Adjust.EndLoadData();}return fixedGroups.Count;
 }
 DataRow NewPayment(string task,Receipt info,string target,decimal? amount){
  var record=Records.FirstOrDefault(r=>r.Task==task);var previous=Payment.AsEnumerable().FirstOrDefault(r=>Val(r,"任务ID")==task);var p=Payment.NewRow();p["任务ID"]=task;p["任务名称"]=record!=null?record.Name:previous==null?"":Val(previous,"任务名称");p["类型(必填)"]=record!=null?record.Type:previous==null?"":Val(previous,"类型(必填)");
  p["收款类型(必填)"]="小额";p["打款对象类型(必填)"]=target;p["团长比例"]=amount.HasValue?(object)DBNull.Value:"1";p["团长打款金额"]=amount.HasValue?(object)amount.Value.ToString(CultureInfo.InvariantCulture):DBNull.Value;
  p["收款平台(必填)"]=info.Channel;p["银行卡号"]=info.Channel=="银行卡"?info.Account:"";p["支付宝账号"]=info.Channel=="支付宝"?info.Account:"";p["身份证号(必填)"]=info.Identity;p["姓名(必填)"]=info.Name;p["手机号(必填)"]=info.Phone;return p;
 }
 void ReplacePayments(IEnumerable<string> tasks,Receipt leader,Receipt creator,string target,decimal? amount){
  var ids=new HashSet<string>(tasks);if(ids.Count==0||ids.Any(id=>String.IsNullOrWhiteSpace(id)||!Records.Any(r=>r.Task==id)&&!Payment.AsEnumerable().Any(r=>Val(r,"任务ID")==id)))throw new Exception("请选择有效任务。");
  ReceiptStore.Validate(leader);if(creator!=null)ReceiptStore.Validate(creator);
  if(creator==null&&!new[]{"团长","达人"}.Contains(target))throw new Exception("请选择团长或达人。");
  if(amount.HasValue&&(amount.Value<0||Decimal.Round(amount.Value,2)!=amount.Value))throw new Exception("金额应为非负数，最多两位小数。");
  // Validate every input before replacing any task. Preserve unrelated task rows.
  var rows=new List<DataRow>();foreach(string id in ids){rows.Add(NewPayment(id,leader,creator==null?target:"团长",amount));if(creator!=null)rows.Add(NewPayment(id,creator,"达人",amount));}
  Payment.BeginLoadData();try{foreach(var r in Payment.AsEnumerable().Where(r=>ids.Contains(Val(r,"任务ID"))).ToList())Payment.Rows.Remove(r);foreach(var r in rows)Payment.Rows.Add(r);}finally{Payment.EndLoadData();}
 }
 public void UseReceipt(Receipt info,string target,IEnumerable<string> tasks=null){ReplacePayments(tasks??Payment.AsEnumerable().Select(r=>Val(r,"任务ID")).Where(id=>id!="").Distinct(),info,null,target,null);}
 public void UseLeaderAmount(Receipt leader,Receipt creator,decimal amount,IEnumerable<string> tasks){ReplacePayments(tasks,leader,creator,"团长",amount);}
}

public sealed class ReceiptRecord {
 public string Id,Label;public Receipt Info;
 public override string ToString(){return Label;}
}
public sealed class ReceiptFile {public int Version=1;public List<ReceiptRecord> Records=new List<ReceiptRecord>();}
public sealed class ReceiptStore {
 public readonly string Path;public List<ReceiptRecord> Records=new List<ReceiptRecord>();
 static readonly JavaScriptSerializer Json=new JavaScriptSerializer{MaxJsonLength=16777216};
 public ReceiptStore(string path){Path=path;Reload();}
 public void Reload(){if(!File.Exists(Path)){Records=new List<ReceiptRecord>();return;}try{var f=Json.Deserialize<ReceiptFile>(File.ReadAllText(Path));if(f==null||f.Version!=1||f.Records==null)throw new Exception("记录文件版本或格式无效");foreach(var r in f.Records){if(r==null||r.Info==null||String.IsNullOrWhiteSpace(r.Id))throw new Exception("收款记录缺少信息");Validate(r.Info);}if(f.Records.Select(r=>r.Id).Distinct().Count()!=f.Records.Count)throw new Exception("收款记录ID重复");Records=f.Records;Relabel();}catch(Exception e){throw new Exception("本地收款记录读取失败，原文件已保留："+Path+"\n"+e.Message);}}
 public static void Validate(Receipt r){if(r==null)throw new Exception("请先选择收款记录。");if(!new[]{"银行卡","支付宝"}.Contains(r.Channel))throw new Exception("请选择收款平台。");if(String.IsNullOrWhiteSpace(r.Name)||String.IsNullOrWhiteSpace(r.Account))throw new Exception("请填写姓名和收款账号。");if(r.Channel=="银行卡"&&!Regex.IsMatch(r.Account,@"^\d{10,30}$"))throw new Exception("银行卡号应为10至30位数字。");if(r.Channel=="支付宝"&&(r.Account.Length<4||r.Account.Any(Char.IsWhiteSpace)))throw new Exception("支付宝账号至少四位且不能包含空格。");if(!Regex.IsMatch(r.Phone??"",@"^1\d{10}$"))throw new Exception("手机号应为11位数字。");if(!Regex.IsMatch(r.Identity??"",@"^(\d{15}|\d{17}[0-9Xx])$"))throw new Exception("请填写有效格式的身份证号。");}
 static string Signature(Receipt r){return String.Join("\u001f",new[]{r.Name,r.Identity.ToUpperInvariant(),r.Phone,r.Channel,r.Account});}
 static string Letters(int index){string s="";for(int n=index+1;n>0;n=(n-1)/26)s=(char)('A'+(n-1)%26)+s;return s;}
 static string Tail(string s){return s.Substring(Math.Max(0,s.Length-4));}
 void Relabel(){foreach(var group in Records.GroupBy(r=>r.Info.Name)){var all=group.ToList();for(int i=0;i<all.Count;i++){var r=all[i];r.Label=r.Info.Name+(all.Count>1?Letters(i):"")+"_"+Tail(r.Info.Phone)+"_"+Tail(r.Info.Account);}}}
 public string Save(Receipt info,string editId=null){Validate(info);var normalized=new Receipt{Channel=info.Channel,Name=info.Name.Trim(),Identity=info.Identity.Trim().ToUpperInvariant(),Phone=info.Phone.Trim(),Account=info.Account.Trim()};var next=Records.Select(r=>new ReceiptRecord{Id=r.Id,Label=r.Label,Info=r.Info}).ToList();var duplicate=next.FirstOrDefault(r=>Signature(r.Info)==Signature(normalized));if(duplicate!=null&&duplicate.Id!=editId){if(editId!=null)throw new Exception("相同收款信息已存在，请直接选择已有记录。");return duplicate.Id;}
  string id=editId??Guid.NewGuid().ToString("N");var existing=next.FirstOrDefault(r=>r.Id==id);if(editId!=null&&existing==null)throw new Exception("编辑的收款记录不存在。");if(existing==null)next.Add(new ReceiptRecord{Id=id,Info=normalized});else existing.Info=normalized;Persist(next);return id;}
 public string Snapshot(){return Json.Serialize(new ReceiptFile{Records=Records});}
 public void Restore(string snapshot){var file=Json.Deserialize<ReceiptFile>(snapshot);if(file==null||file.Version!=1||file.Records==null)throw new Exception("收款记录快照无效。");foreach(var r in file.Records)Validate(r.Info);Persist(file.Records);}
 public void Delete(string id){Persist(Records.Where(r=>r.Id!=id).ToList());}
 void Persist(List<ReceiptRecord> next){var old=Records;Records=next;Relabel();string tmp=Path+"."+Guid.NewGuid().ToString("N")+".tmp";try{Directory.CreateDirectory(System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(Path)));File.WriteAllText(tmp,Json.Serialize(new ReceiptFile{Records=Records}),new UTF8Encoding(false));if(File.Exists(Path))File.Replace(tmp,Path,Path+".bak",true);else File.Move(tmp,Path);}catch{Records=old;Relabel();throw;}finally{if(File.Exists(tmp))File.Delete(tmp);}}
}

// Independent view state: filtering never removes model rows or changes export scope.
public sealed class GridFilter {
 sealed class FilterValue {public string Value;public override string ToString(){return Value==""?"（空白）":Value;}}
 readonly DataGridView Grid;DataTable Table;readonly Dictionary<string,HashSet<string>> Filters=new Dictionary<string,HashSet<string>>();
 public readonly Panel Host;public readonly TextBox Search;readonly ComboBox Column;readonly Label Count;
 public GridFilter(DataGridView grid){Grid=grid;Host=new FlowLayoutPanel{Dock=DockStyle.Top,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,WrapContents=true,Padding=Padding.Empty,Margin=Padding.Empty};Column=new ComboBox{Left=0,Top=5,Width=128,DropDownStyle=ComboBoxStyle.DropDownList};var choose=new Button{Text="筛选 ▼",Left=134,Top=4,Width=80,Height=30};var clear=new Button{Text="清除筛选",Left=220,Top=4,Width=88,Height=30};Search=new TextBox{Left=362,Top=6,Width=180};var label=new Label{Text="查找",Left=320,Top=10,Width=40};Count=new Label{Left=550,Top=10,AutoSize=true,ForeColor=Color.FromArgb(95,110,130)};foreach(Control control in new Control[]{Column,choose,clear,label,Search,Count}){control.Margin=new Padding(0,0,8,0);if(control is Label){control.AutoSize=true;control.Padding=new Padding(0,4,0,0);}}Host.Controls.AddRange(new Control[]{Column,choose,clear,label,Search,Count});choose.Click+=(o,e)=>{if(Column.SelectedItem!=null)Choose(Convert.ToString(Column.SelectedItem));};clear.Click+=(o,e)=>Clear();Search.TextChanged+=(o,e)=>Apply();Grid.ColumnHeaderMouseClick+=(o,e)=>{if(e.Button==MouseButtons.Right&&e.ColumnIndex>=0)Choose(Grid.Columns[e.ColumnIndex].DataPropertyName);};Grid.KeyDown+=(o,e)=>{if(e.Control&&e.KeyCode==Keys.F){Search.Focus();Search.SelectAll();e.Handled=true;}};}
 public void Bind(DataTable table){Table=table;Filters.Clear();Search.Clear();Grid.DataSource=table.DefaultView;Column.Items.Clear();foreach(DataColumn c in table.Columns){var gc=Grid.Columns[c.ColumnName];if(gc!=null&&gc.Visible)Column.Items.Add(c.ColumnName);}if(Column.Items.Count>0)Column.SelectedIndex=0;Apply();}
 static string Field(string name){return "["+name.Replace("\\","\\\\").Replace("]","\\]")+ "]";}
 static string Literal(string text){return "'"+text.Replace("'","''")+"'";}
 static string Like(string text){var b=new StringBuilder();foreach(char c in text){if(c=='[')b.Append("[[]");else if(c==']')b.Append("[]]");else if(c=='%')b.Append("[%]");else if(c=='*')b.Append("[*]");else b.Append(c);}return b.ToString();}
 public void SelectValues(string column,IEnumerable<string> values){if(Table==null)return;if(!Table.Columns.Contains(column))throw new Exception("筛选列不存在。");Filters[column]=new HashSet<string>(values);Apply();}
 public void Clear(){Filters.Clear();Search.Clear();Apply();}
 public void Apply(){if(Table==null)return;var clauses=new List<string>();foreach(var f in Filters){string field="ISNULL(CONVERT("+Field(f.Key)+", 'System.String'), '')";clauses.Add(f.Value.Count==0?"1 = 0":"("+String.Join(" OR ",f.Value.Select(v=>field+" = "+Literal(v)))+")");}if(Search.Text!="")clauses.Add("("+String.Join(" OR ",Table.Columns.Cast<DataColumn>().Select(c=>"ISNULL(CONVERT("+Field(c.ColumnName)+", 'System.String'), '') LIKE "+Literal("%"+Like(Search.Text)+"%")))+")");Table.DefaultView.RowFilter=String.Join(" AND ",clauses);Count.Text=Table.DefaultView.Count+" / "+Table.Rows.Count+" 行";}
 void Choose(string column){if(Table==null)return;Grid.EndEdit();var values=Table.AsEnumerable().Select(r=>Model.Val(r,column)).Distinct().OrderBy(x=>x,StringComparer.CurrentCulture).ToList();HashSet<string> previous;var selected=Filters.TryGetValue(column,out previous)?new HashSet<string>(previous):new HashSet<string>(values);
  using(var f=new Form{Text="筛选："+column,Size=new Size(410,480),MinimumSize=new Size(410,480),StartPosition=FormStartPosition.CenterParent,Font=Grid.Font}){var hint=new Label{Text="勾选要显示的值，可搜索后选择",Dock=DockStyle.Top,Height=30};var query=new TextBox{Dock=DockStyle.Top};var list=new CheckedListBox{Dock=DockStyle.Fill,CheckOnClick=true};var bottom=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=45};bool loading=false;Action populate=()=>{loading=true;list.Items.Clear();foreach(string v in values.Where(v=>v.IndexOf(query.Text,StringComparison.CurrentCultureIgnoreCase)>=0))list.Items.Add(new FilterValue{Value=v},selected.Contains(v));loading=false;};list.ItemCheck+=(o,e)=>{if(loading)return;string v=((FilterValue)list.Items[e.Index]).Value;if(e.NewValue==CheckState.Checked)selected.Add(v);else selected.Remove(v);};query.TextChanged+=(o,e)=>populate();var all=new Button{Text="全选",AutoSize=true};var none=new Button{Text="全不选",AutoSize=true};var ok=new Button{Text="应用",AutoSize=true,DialogResult=DialogResult.OK};var cancel=new Button{Text="取消",AutoSize=true,DialogResult=DialogResult.Cancel};all.Click+=(o,e)=>{foreach(string v in values.Where(v=>v.IndexOf(query.Text,StringComparison.CurrentCultureIgnoreCase)>=0))selected.Add(v);populate();};none.Click+=(o,e)=>{foreach(string v in values.Where(v=>v.IndexOf(query.Text,StringComparison.CurrentCultureIgnoreCase)>=0))selected.Remove(v);populate();};bottom.Controls.AddRange(new Control[]{all,none,ok,cancel});f.Controls.Add(list);f.Controls.Add(query);f.Controls.Add(hint);f.Controls.Add(bottom);populate();if(f.ShowDialog(Grid.FindForm())==DialogResult.OK)SelectValues(column,selected);}
 }
}

public sealed class ReceiptEditor:Form {
 readonly ComboBox channel;readonly TextBox account,identity,name,phone;public Receipt Result;
 public ReceiptEditor(Receipt info){Text=info==null?"新增收款记录":"编辑收款记录";Size=new Size(440,465);FormBorderStyle=FormBorderStyle.FixedDialog;MaximizeBox=false;MinimizeBox=false;StartPosition=FormStartPosition.CenterParent;Font=new Font("Microsoft YaHei UI",10);var layout=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=2,Padding=new Padding(20),RowCount=7};layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,100));layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));Controls.Add(layout);channel=new ComboBox{DropDownStyle=ComboBoxStyle.DropDownList,Dock=DockStyle.Top};channel.Items.AddRange(new[]{"银行卡","支付宝"});channel.SelectedIndex=info!=null&&info.Channel=="支付宝"?1:0;account=new TextBox{Text=info==null?"":info.Account,Dock=DockStyle.Top};identity=new TextBox{Text=info==null?"":info.Identity,Dock=DockStyle.Top};name=new TextBox{Text=info==null?"":info.Name,Dock=DockStyle.Top};phone=new TextBox{Text=info==null?"":info.Phone,Dock=DockStyle.Top};var fields=new Control[]{channel,account,identity,name,phone};string[] names={"收款平台","收款账号","身份证号","姓名","手机号"};for(int i=0;i<5;i++){layout.RowStyles.Add(new RowStyle(SizeType.Absolute,53));layout.Controls.Add(new Label{Text=names[i],AutoSize=true,Padding=new Padding(0,6,0,0)},0,i);layout.Controls.Add(fields[i],1,i);}var note=new Label{Text="保存后可重复选用；账号按文本保存。",Dock=DockStyle.Fill};layout.Controls.Add(note,0,5);layout.SetColumnSpan(note,2);var save=new Button{Text="保存记录",AutoSize=true};save.Click+=(o,e)=>{try{var r=new Receipt{Channel=Convert.ToString(channel.SelectedItem),Account=account.Text.Trim(),Identity=identity.Text.Trim(),Name=name.Text.Trim(),Phone=phone.Text.Trim()};ReceiptStore.Validate(r);Result=r;DialogResult=DialogResult.OK;}catch(Exception ex){MessageBox.Show(this,ex.Message,"请检查");}};layout.Controls.Add(save,1,6);AcceptButton=save;}
}

public sealed class LeaderAmountWindow:Form {
 readonly Model Model;readonly ComboBox leader,creator;readonly TextBox amount;readonly CheckedListBox tasks;public int Applied;
 public LeaderAmountWindow(Model model,IEnumerable<ReceiptRecord> records){Model=model;Text="团长打款金额";Size=new Size(620,585);MinimumSize=Size;StartPosition=FormStartPosition.CenterParent;Font=new Font("Microsoft YaHei UI",10);var layout=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=2,RowCount=6,Padding=new Padding(20)};layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,115));layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));Controls.Add(layout);leader=new ComboBox{Dock=DockStyle.Top,DropDownStyle=ComboBoxStyle.DropDownList};creator=new ComboBox{Dock=DockStyle.Top,DropDownStyle=ComboBoxStyle.DropDownList};var choices=records.ToArray();leader.Items.AddRange(choices);creator.Items.AddRange(choices);amount=new TextBox{Dock=DockStyle.Top};tasks=new CheckedListBox{Dock=DockStyle.Fill,CheckOnClick=true};foreach(string id in model.Payment.AsEnumerable().Select(r=>Model.Val(r,"任务ID")).Where(id=>id!="").Distinct())tasks.Items.Add(id,true);Control[] fields={leader,creator,amount,tasks};string[] labels={"团长收款记录","达人收款记录","团长打款金额","对应任务"};for(int i=0;i<4;i++){layout.RowStyles.Add(new RowStyle(i==3?SizeType.Percent:SizeType.Absolute,i==3?100:50));layout.Controls.Add(new Label{Text=labels[i],AutoSize=true,Padding=new Padding(0,6,0,0)},0,i);layout.Controls.Add(fields[i],1,i);}layout.RowStyles.Add(new RowStyle(SizeType.Absolute,65));var note=new Label{Text="每个勾选任务生成团长、达人两行，两行金额一致，团长比例留空。再次录入会替换这些任务的打款记录。",Dock=DockStyle.Fill};layout.Controls.Add(note,0,4);layout.SetColumnSpan(note,2);layout.RowStyles.Add(new RowStyle(SizeType.Absolute,42));var apply=new Button{Text="应用并预览",AutoSize=true};apply.Click+=(o,e)=>{try{decimal n;if(!Decimal.TryParse(amount.Text,NumberStyles.Number,CultureInfo.CurrentCulture,out n))throw new Exception("请输入有效的团长打款金额。");var l=leader.SelectedItem as ReceiptRecord;var c=creator.SelectedItem as ReceiptRecord;if(l==null||c==null)throw new Exception("请分别选择团长和达人的收款记录。");var ids=tasks.CheckedItems.Cast<string>().ToList();Model.UseLeaderAmount(l.Info,c.Info,n,ids);Applied=ids.Count;DialogResult=DialogResult.OK;}catch(Exception ex){MessageBox.Show(this,ex.Message,"请检查");}};layout.Controls.Add(apply,1,5);AcceptButton=apply;}
}
}
