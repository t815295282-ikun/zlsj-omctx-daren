using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace CreatorSettlement {
public sealed partial class MainWindow {
 sealed class EditState {
  public DataTable Check,Payment,Adjust,Warnings;
  public string Receipts,Selected;
  public int Page,Target;
 }
 sealed class EditStep {public EditState Before,After;public string Name;}
 readonly List<EditStep> undoSteps=new List<EditStep>(),redoSteps=new List<EditStep>();
 EditState cellBefore;string cellName;bool restoring,committing;
 Button undoButton,redoButton;
 EditState CaptureState(){var selected=receiptList.SelectedItem as ReceiptRecord;return new EditState{Check=model==null?null:model.Check.Copy(),Payment=model==null?null:model.Payment.Copy(),Adjust=model==null?null:model.Adjust.Copy(),Warnings=model==null?null:model.Warnings.Copy(),Receipts=receiptStore.Snapshot(),Selected=selected==null?null:selected.Id,Page=Math.Max(0,tabs.SelectedIndex),Target=paymentTarget.SelectedIndex};}
 static bool SameTable(DataTable a,DataTable b){if(a==null||b==null)return a==b;if(a.Rows.Count!=b.Rows.Count)return false;for(int i=0;i<a.Rows.Count;i++)for(int j=0;j<a.Columns.Count;j++)if(!Object.Equals(a.Rows[i][j],b.Rows[i][j]))return false;return true;}
 static bool SameState(EditState a,EditState b){return a.Receipts==b.Receipts&&SameTable(a.Check,b.Check)&&SameTable(a.Payment,b.Payment)&&SameTable(a.Adjust,b.Adjust)&&SameTable(a.Warnings,b.Warnings);}
 static void RestoreTable(DataTable target,DataTable saved){target.BeginLoadData();try{target.Clear();foreach(DataRow row in saved.Rows)target.ImportRow(row);}finally{target.EndLoadData();}}
 void RestoreState(EditState saved){restoring=true;try{if(receiptStore.Snapshot()!=saved.Receipts)receiptStore.Restore(saved.Receipts);if(model!=null&&saved.Check!=null){RestoreTable(model.Check,saved.Check);RestoreTable(model.Payment,saved.Payment);RestoreTable(model.Adjust,saved.Adjust);RestoreTable(model.Warnings,saved.Warnings);model.Refresh();checkFilter.Apply();payFilter.Apply();adjustFilter.Apply();}RefreshReceipts(saved.Selected);paymentTarget.SelectedIndex=saved.Target;tabs.SelectedIndex=saved.Page;UpdateMetrics();UpdateNavigation();}finally{restoring=false;}}
 void RecordStep(EditState before,string name){if(restoring)return;var after=CaptureState();if(SameState(before,after))return;undoSteps.Add(new EditStep{Before=before,After=after,Name=name});if(undoSteps.Count>50)undoSteps.RemoveAt(0);redoSteps.Clear();UpdateHistoryButtons();}
 void RunChange(string name,Action action){if(model!=null)Commit();var before=CaptureState();try{action();RecordStep(before,name);}catch{RestoreState(before);throw;}}
 void ResetHistory(){undoSteps.Clear();redoSteps.Clear();cellBefore=null;UpdateHistoryButtons();}
 void UpdateHistoryButtons(){if(undoButton==null)return;undoButton.Enabled=undoSteps.Count>0||cellBefore!=null;redoButton.Enabled=redoSteps.Count>0;}
 void TrackGrid(DataGridView grid){grid.CellBeginEdit+=(o,e)=>{if(restoring)return;cellBefore=CaptureState();cellName="修改"+(grid==checkGrid?"核对表":grid==payGrid?"打款表":grid==warningGrid?"异常表":"调差表");UpdateHistoryButtons();};grid.CellEndEdit+=(o,e)=>{if(restoring||cellBefore==null)return;var before=cellBefore;cellBefore=null;Attempt(()=>{try{BindingContext[(grid==checkGrid?model.Check:grid==payGrid?model.Payment:grid==warningGrid?model.Warnings:model.Adjust).DefaultView].EndCurrentEdit();model.Refresh();Commit();RecordStep(before,cellName);}catch{RestoreState(before);UpdateHistoryButtons();throw;}});UpdateHistoryButtons();};}
 void Undo(){if(model!=null)Commit();if(undoSteps.Count==0)return;var step=undoSteps[undoSteps.Count-1];RestoreState(step.Before);undoSteps.RemoveAt(undoSteps.Count-1);redoSteps.Add(step);UpdateHistoryButtons();status.Text="已撤销："+step.Name;}
 void Redo(){if(model!=null)Commit();if(redoSteps.Count==0)return;var step=redoSteps[redoSteps.Count-1];RestoreState(step.After);redoSteps.RemoveAt(redoSteps.Count-1);undoSteps.Add(step);UpdateHistoryButtons();status.Text="已重做："+step.Name;}
 protected override bool ProcessCmdKey(ref Message msg,Keys keyData){if(keyData==(Keys.Control|Keys.Z)){Attempt(Undo);return true;}if(keyData==(Keys.Control|Keys.Y)||keyData==(Keys.Control|Keys.Shift|Keys.Z)){Attempt(Redo);return true;}if(ExcelShortcut(keyData))return true;return base.ProcessCmdKey(ref msg,keyData);}
}
}
