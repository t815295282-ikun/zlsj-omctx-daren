using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace CreatorSettlement {
public sealed partial class MainWindow {
 DataGridView CurrentGrid(){return new[]{checkGrid,payGrid,warningGrid,adjustGrid}[Math.Max(0,tabs.SelectedIndex)];}
 DataTable CurrentTable(){return new[]{model.Check,model.Payment,model.Warnings,model.Adjust}[Math.Max(0,tabs.SelectedIndex)];}
 void AddTableRow(){Require();Commit();var table=CurrentTable();var row=table.NewRow();foreach(DataColumn column in table.Columns)if(column.Expression=="")row[column]=column.DataType==typeof(decimal)?(object)DBNull.Value:"";if(table==model.Payment){row["收款类型(必填)"]="小额";row["打款对象类型(必填)"]="团长";row["团长比例"]="1";}using(var editor=new TableRowEditor(row,table==model.Check)){if(editor.ShowDialog(this)!=DialogResult.OK)return;RunChange("新增一行",()=>{editor.Apply(row);table.Rows.Add(row);model.Refresh();});}CurrentFilterClear();status.Text="已新增一行；可以继续修改，也可以撤销。";}
 void EditTableRow(){Require();Commit();var grid=CurrentGrid();if(grid.CurrentRow==null)throw new Exception("请先点击要修改的一行。");var row=((DataRowView)grid.CurrentRow.DataBoundItem).Row;using(var editor=new TableRowEditor(row,CurrentTable()==model.Check)){if(editor.ShowDialog(this)!=DialogResult.OK)return;RunChange("修改选中行",()=>{editor.Apply(row);model.Refresh();});}RefreshFilters();status.Text="已修改选中行；可以撤销恢复。";}
 void DeleteTableRows(){Require();Commit();var grid=CurrentGrid();var rows=grid.SelectedCells.Cast<DataGridViewCell>().Select(c=>((DataRowView)c.OwningRow.DataBoundItem).Row).Distinct().ToList();if(rows.Count==0)throw new Exception("请先选择要删除的行中的单元格。");RunChange("删除选中行",()=>{foreach(var row in rows)row.Table.Rows.Remove(row);model.Refresh();});RefreshFilters();UpdateMetrics();status.Text="已删除 "+rows.Count+" 行；点击撤销可恢复。";}
 void RefreshFilters(){checkFilter.Apply();payFilter.Apply();adjustFilter.Apply();}
 void CurrentFilterClear(){int i=tabs.SelectedIndex;if(i==0)checkFilter.Clear();if(i==1)payFilter.Clear();if(i==3)adjustFilter.Clear();UpdateMetrics();}
}
public sealed class TableRowEditor:Form {
 readonly Dictionary<DataColumn,TextBox> inputs=new Dictionary<DataColumn,TextBox>();
 readonly Dictionary<DataColumn,object> values=new Dictionary<DataColumn,object>();
 public TableRowEditor(DataRow row,bool check){Text="填写这一行";Size=new Size(590,660);MinimumSize=new Size(500,400);StartPosition=FormStartPosition.CenterParent;Font=new Font("Microsoft YaHei UI",10);var body=new Panel{Dock=DockStyle.Fill,AutoScroll=true,Padding=new Padding(16)};var form=new TableLayoutPanel{Dock=DockStyle.Top,AutoSize=true,ColumnCount=2,Padding=new Padding(0,0,12,0)};form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,185));form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));int index=0;foreach(DataColumn column in row.Table.Columns){if(column.Expression!=""||check&&(column.Ordinal==6||column.Ordinal==8))continue;var input=new TextBox{Dock=DockStyle.Fill,Text=row.IsNull(column)?"":Convert.ToString(row[column],CultureInfo.CurrentCulture)};inputs[column]=input;form.RowStyles.Add(new RowStyle(SizeType.Absolute,43));form.Controls.Add(new Label{Text=column.ColumnName,AutoSize=true,Padding=new Padding(0,6,0,0)},0,index);form.Controls.Add(input,1,index++);}body.Controls.Add(form);var buttons=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=60,Padding=new Padding(18,10,0,0)};var save=new Button{Text="保存这一行",AutoSize=true};var cancel=new Button{Text="取消",AutoSize=true,DialogResult=DialogResult.Cancel};buttons.Controls.AddRange(new Control[]{save,cancel});Controls.Add(body);Controls.Add(buttons);AcceptButton=save;CancelButton=cancel;save.Click+=(o,e)=>{values.Clear();foreach(var pair in inputs){string text=pair.Value.Text.Trim();if(pair.Key.DataType==typeof(decimal)){decimal number;if(text!=""&&!Decimal.TryParse(text,NumberStyles.Number,CultureInfo.CurrentCulture,out number)){MessageBox.Show(this,"“"+pair.Key.ColumnName+"”请填写有效数字。");pair.Value.Focus();return;}values[pair.Key]=text==""?(object)DBNull.Value:Decimal.Parse(text,CultureInfo.CurrentCulture);}else values[pair.Key]=text;}if(row.Table.Columns.Contains("任务ID")&&Convert.ToString(values[row.Table.Columns["任务ID"]])==""||check&&Convert.ToString(values[row.Table.Columns["任务/方案ID"]])==""){MessageBox.Show(this,"请填写任务ID，方便核对和同步。");return;}if(check&&values[row.Table.Columns["一口价（无手续费）"]]!=DBNull.Value&&(decimal)values[row.Table.Columns["一口价（无手续费）"]]<0){MessageBox.Show(this,"一口价成本不能小于0，实际打款可以为负数。");return;}DialogResult=DialogResult.OK;};}
 public void Apply(DataRow row){foreach(var pair in values)row[pair.Key.ColumnName]=pair.Value;}
}
}
