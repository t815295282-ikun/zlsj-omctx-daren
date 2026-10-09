using System;
using System.Drawing;
using System.Windows.Forms;
namespace CreatorSettlement {
public sealed partial class MainWindow {
 void ConfigureDpi(){AutoScaleDimensions=new SizeF(96,96);AutoScaleMode=AutoScaleMode.Dpi;Shown+=(o,e)=>FitScreen();DpiChanged+=(o,e)=>{if(IsHandleCreated)BeginInvoke((Action)FitScreen);};}
 void FitScreen(){var area=Screen.FromControl(this).WorkingArea;int maxW=Math.Max(400,area.Width-24),maxH=Math.Max(300,area.Height-24);MinimumSize=new Size(Math.Min(MinimumSize.Width,maxW),Math.Min(MinimumSize.Height,maxH));if(WindowState==FormWindowState.Normal){Size=new Size(Math.Min(Width,maxW),Math.Min(Height,maxH));Location=new Point(Math.Max(area.Left,Math.Min(Left,area.Right-Width)),Math.Max(area.Top,Math.Min(Top,area.Bottom-Height)));}foreach(Control control in Controls)if(control.Dock==DockStyle.Fill)control.Bounds=ClientRectangle;LayoutChildren(this);}
 static void LayoutChildren(Control parent){parent.PerformLayout();foreach(Control child in parent.Controls)LayoutChildren(child);}
}
}
