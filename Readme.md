<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128540062/13.1.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/E3626)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# ASP.NET Web Forms - How to export several controls to different XLSX worksheets

This example illustrates how to export several components to different worksheets of the same XLSX document.

![Exported Document](exported-document.gif)

In this example, the page contains 3 components: [ASPxGridView](https://docs.devexpress.com/AspNet/DevExpress.Web.ASPxGridView), [ASPxTreeList](https://docs.devexpress.com/AspNet/DevExpress.Web.ASPxTreeList.ASPxTreeList), and [WebChartControl](https://docs.devexpress.com/AspNet/DevExpress.XtraCharts.Web.WebChartControl). You can export them to different worksheets of the same document in the following way:

1. Create a [PrintableComponentLinkBase](https://docs.devexpress.com/CoreLibraries/DevExpress.XtraPrintingLinks.PrintableComponentLinkBase) object for every component.

    ```cs
    PrintableComponentLinkBase link1 = new PrintableComponentLinkBase(ps);
    link1.Component = Grid;

    PrintableComponentLinkBase link2 = new PrintableComponentLinkBase(ps);
    link2.Component = Tree;

    PrintableComponentLinkBase link3 = new PrintableComponentLinkBase(ps);
    Chart.DataBind();
    link3.Component = ((IChartContainer)Chart).Chart;
    ```

2. Call the [CreatePageForEachLink](https://docs.devexpress.com/CoreLibraries/DevExpress.XtraPrintingLinks.CompositeLinkBase.CreatePageForEachLink) method to create a separate page for every component.

    ```cs
    CompositeLinkBase compositeLink = new CompositeLinkBase(ps);
    compositeLink.Links.AddRange(new object[] { link1, link2, link3 });
    compositeLink.CreatePageForEachLink();
    ```

3. Create an [XlsxExportOptions](https://docs.devexpress.com/CoreLibraries/DevExpress.XtraPrinting.XlsExportOptions) object and set its [ExportMode](https://docs.devexpress.com/CoreLibraries/DevExpress.XtraPrinting.XlsExportOptions.ExportMode) property to `SingleFilePageByPage` value to export a document to a single file, page-by-page. Send the options to the [ExportToXlsx](https://docs.devexpress.com/CoreLibraries/DevExpress.XtraPrinting.PrintingSystemBase.ExportToXlsx(System.IO.Stream-DevExpress.XtraPrinting.XlsxExportOptions)) method to export the document.

    ```cs
    XlsxExportOptions options = new XlsxExportOptions();
    options.ExportMode = XlsxExportMode.SingleFilePageByPage;
    compositeLink.PrintingSystemBase.ExportToXlsx(stream, options);
    ```


## Files to Review

* [Default.aspx](./CS/WebSite/Default.aspx) (VB: [Default.aspx](./VB/WebSite/Default.aspx))
* [Default.aspx.cs](./CS/WebSite/Default.aspx.cs) (VB: [Default.aspx.vb](./VB/WebSite/Default.aspx.vb))

## More Examples

* [How to export ASPxPivotGrid and bound WebChartControl to the same print document](https://github.com/DevExpress-Examples/asp-net-web-forms-export-pivot-grid-and-chart-in-one-document)
* [How to export ASPxGridView and WebChartControl to the same print document](https://github.com/DevExpress-Examples/asp-net-web-forms-export-grid-and-chart-in-one-document)
* [Grid View for Web Forms - How to export multiple grids into a single print document](https://github.com/DevExpress-Examples/asp-net-web-forms-grid-export-two-grids-in-one-document)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=asp-net-web-forms-export-several-controls-to-different-sheets&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=asp-net-web-forms-export-several-controls-to-different-sheets&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
