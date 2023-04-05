using System;
using System.IO;
using DevExpress.XtraCharts;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrintingLinks;

public partial class _Default : System.Web.UI.Page {
    protected void Page_Load(object sender, EventArgs e) {
        if(!IsPostBack) {
            Tree.DataBind();
            Tree.ExpandToLevel(2);
        }

        Chart.SeriesDataMember = "CategoryName";
        Chart.SeriesTemplate.ArgumentDataMember = "ProductSales";
        Chart.SeriesTemplate.ValueDataMembers.AddRange(new string[] { "ProductSales" });
        Chart.SeriesTemplate.View = new StackedBarSeriesView();
    }
    protected void ExportButton_Click(object sender, EventArgs e) {
        PrintingSystemBase ps = new PrintingSystemBase();

        PrintableComponentLinkBase link1 = new PrintableComponentLinkBase(ps);
        link1.Component = Grid;

        PrintableComponentLinkBase link2 = new PrintableComponentLinkBase(ps);
        link2.Component = Tree;

        PrintableComponentLinkBase link3 = new PrintableComponentLinkBase(ps);
        Chart.DataBind();
        link3.Component = ((IChartContainer)Chart).Chart;

        CompositeLinkBase compositeLink = new CompositeLinkBase(ps);
        compositeLink.Links.AddRange(new object[] { link1, link2, link3 });

        compositeLink.CreatePageForEachLink();

        using(MemoryStream stream = new MemoryStream()) {
            XlsxExportOptions options = new XlsxExportOptions();
            options.ExportMode = XlsxExportMode.SingleFilePageByPage;
            compositeLink.PrintingSystemBase.ExportToXlsx(stream, options);
            Response.Clear();
            Response.Buffer = false;
            Response.AppendHeader("Content-Type", "application/xlsx");
            Response.AppendHeader("Content-Transfer-Encoding", "binary");
            Response.AppendHeader("Content-Disposition", "attachment; filename=test.xlsx");
            Response.BinaryWrite(stream.ToArray());
            Response.End();
        }
        ps.Dispose();
    }
}
