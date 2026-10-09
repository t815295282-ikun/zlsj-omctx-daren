using System;
using System.Linq;
using System.Windows.Forms;

namespace CreatorSettlement {
public static class GridCopy {
 public static string SelectionText(DataGridView grid){var data=grid.GetClipboardContent();return data==null?"":Convert.ToString(data.GetData(DataFormats.UnicodeText));}
 static string Quote(string value){return value.IndexOfAny(new[]{'\t','\r','\n','"'})<0?value:"\""+value.Replace("\"","\"\"")+"\"";}
 public static string ColumnText(DataGridView grid){if(grid.CurrentCell==null)return "";int col=grid.CurrentCell.ColumnIndex;return String.Join(Environment.NewLine,grid.Rows.Cast<DataGridViewRow>().Where(r=>r.Visible&&!r.IsNewRow).Select(r=>Quote(Convert.ToString(r.Cells[col].FormattedValue))));}
 public static void SelectRightClick(DataGridView grid,int row,int column){if(row<0||column<0)return;var cell=grid.Rows[row].Cells[column];if(!cell.Selected){grid.CurrentCell=cell;grid.ClearSelection();cell.Selected=true;}}
 public static void Attach(DataGridView grid,Action<string> copy,Action<Action> attempt){
  grid.SelectionMode=DataGridViewSelectionMode.CellSelect;grid.MultiSelect=true;grid.ClipboardCopyMode=DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
  var menu=new ContextMenuStrip();var selected=new ToolStripMenuItem("复制选中内容"){ShortcutKeyDisplayString="Ctrl+C"};var column=new ToolStripMenuItem("复制此列（当前显示行）");menu.Items.AddRange(new ToolStripItem[]{selected,column});grid.ContextMenuStrip=menu;
  grid.CellMouseDown+=(o,e)=>{if(e.Button==MouseButtons.Right)SelectRightClick(grid,e.RowIndex,e.ColumnIndex);};
  menu.Opening+=(o,e)=>{selected.Enabled=grid.SelectedCells.Count>0;column.Enabled=grid.CurrentCell!=null&&grid.Rows.Count>0;};
  selected.Click+=(o,e)=>attempt(()=>{if(!grid.EndEdit())throw new Exception("请先修正正在填写的单元格。");string text=SelectionText(grid);if(text!="")copy(text);});
  column.Click+=(o,e)=>attempt(()=>{if(!grid.EndEdit())throw new Exception("请先修正正在填写的单元格。");string text=ColumnText(grid);if(text!="")copy(text);});
  grid.Disposed+=(o,e)=>menu.Dispose();
 }
}
}
