using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CreatorSettlement {
public sealed class SavedFilesWindow:Form {
 readonly string[] files;readonly CheckedListBox list;readonly Button open;
 public string[] SelectedFiles {get{return list.CheckedIndices.Cast<int>().Select(i=>files[i]).ToArray();}}
 public SavedFilesWindow(IEnumerable<string> paths){files=paths.ToArray();if(files.Length==0)throw new ArgumentException("没有已保存的表格。");Text="表格已保存";Size=new Size(670,370);MinimumSize=Size;StartPosition=FormStartPosition.CenterParent;Font=new Font("Microsoft YaHei UI",10);BackColor=Color.White;
  var layout=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=4,Padding=new Padding(18)};layout.RowStyles.Add(new RowStyle(SizeType.Absolute,42));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,50));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.Absolute,58));Controls.Add(layout);
  layout.Controls.Add(new Label{Text="保存好了，要现在打开表格吗？",Dock=DockStyle.Fill,Font=new Font(Font.FontFamily,14,FontStyle.Bold)},0,0);
  layout.Controls.Add(new Label{Text="保存位置："+Path.GetDirectoryName(files[0]),Dock=DockStyle.Fill,AutoEllipsis=true},0,1);
  list=new CheckedListBox{Dock=DockStyle.Fill,CheckOnClick=true,HorizontalScrollbar=true,IntegralHeight=false};foreach(string path in files)list.Items.Add(Path.GetFileName(path),true);layout.Controls.Add(list,0,2);
  var buttons=new FlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft,Padding=new Padding(0,7,0,0)};var later=new Button{Text="暂不打开",Width=112,Height=32,DialogResult=DialogResult.Cancel};open=new Button{Text="立即打开",Width=112,Height=32,DialogResult=DialogResult.OK};buttons.Controls.Add(later);buttons.Controls.Add(open);layout.Controls.Add(buttons,0,3);AcceptButton=open;CancelButton=later;
  list.ItemCheck+=(o,e)=>{int count=list.CheckedItems.Count+(e.NewValue==CheckState.Checked?1:0)-(e.CurrentValue==CheckState.Checked?1:0);open.Enabled=count>0;};
 }
}
public static class SavedFilesPrompt {
 public static List<string> OpenFiles(IEnumerable<string> files,Action<string> opener){var failures=new List<string>();foreach(string path in files){try{opener(path);}catch(Exception ex){failures.Add(Path.GetFileName(path)+"："+ex.Message);}}return failures;}
 public static void Show(Form owner,string[] files,Action<string> opener){using(var f=new SavedFilesWindow(files)){if(f.ShowDialog(owner)!=DialogResult.OK)return;var errors=OpenFiles(f.SelectedFiles,opener);if(errors.Count>0)MessageBox.Show(owner,"表格已保存，但以下文件没能打开：\n"+String.Join("\n",errors)+"\n\n可以到保存位置手动打开，或检查电脑是否安装了 Excel / WPS。","表格已保存",MessageBoxButtons.OK,MessageBoxIcon.Information);}}
}
}
