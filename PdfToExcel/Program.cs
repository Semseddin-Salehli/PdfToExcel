using Microsoft.AspNetCore.Http.HttpResults;
using Spire.Pdf;
using Spire.Pdf.Conversion;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

// Create a PdfDocument instance
PdfDocument document = new PdfDocument();
String path = "C:\\Users\\HP\\Desktop\\gamaproje.pdf";


//Load a PDF file
document.LoadFromFile(@path);

//The four parameters represent: convertToMultipleSheet, showRotatedText, splitCell, wrapText
XlsxLineLayoutOptions options = new XlsxLineLayoutOptions(false, true, true ,true);

//Set PDF to XLSX convert options
document.ConvertOptions.SetPdfToXlsxOptions(options);

//Save to Excel
document.SaveToFile(@path, FileFormat.XLSX);

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
