using TextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;  
using System.ComponentModel;  
using System.Data;  
using System.Drawing;  
using System.Linq;  
using System.Text;  
using System.Windows.Forms;  
  
using System.Data.SqlClient;  

class TextExtractor
{
    PDFLoader loader = new PDFLoader();
    TextParser parser = new TextParser();
    
    private string textToParse = loader.GetTextFromPDF();
    
    public string getCR()
    {
        public string cr = parser.getBetween(textToParse, "#", "report");
        
        return cr;
    }
    
    public string getDate() 
    {
        public string date = parser.getBetween(textToParse, ":", "location");
        
        return date;
    }
    
    public string getOfficer() 
    {
        public string officer = parser.getBetween(textToParse, ":", "date");
        
        return officer;
    }
    
    public string getReportType() 
    {
        public string reportType = parser.getBetween(textToParse, "type", "reporting");
        
        return reportType;
    }
    
    public string getVictim() 
    {
        public string victim = parser.getBetween(textToParse, "", "");
        
        return victim;
    }
    
    public string getLocation() 
    {
        public string location = parser.getBetween(textToParse, ":", "time");
        
        return location;
    }
    
    public string getReportInfo()
    {
        public string report = parser.getBetween(textToParse, "*", "*");
        
        return report;
    }
}

class PDFLoader
{
    private string GetTextFromPDF()
    {
        StringBuilder text = new StringBuilder();
        using (PdfReader reader = new PdfReader("D:\\RentReceiptFormat.pdf"))
        {
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
            }
        }
        
        return text.ToString().ToLower();
    }
}

class TextParser
{
    public string getBetween(string strSource, string strStart, string strEnd)
    {
        if (source.Contains(strStart) && source.Contains(strEnd))
        {
            int Start, End;
            Start = strSource.IndexOf(strStart, 0) + strStart.Length;
            End = strSource.IndexOf(strEnd, Start);
            return strSource.Substring(Start, End - Start);
        }
        
        return "";
    }
}

class InfoHandler
{
    public Dictionary<string, string> info = new Dictionary<string, string>();
    TextExtractor extractor = new TextExtractor();
    
    public InfoHandler()
    {
        info.Add("cr", extractor.getCR());
        info.Add("date", extractor.getDate());
        info.Add("officer", extractor.getOfficer());
        info.Add("type", extractor.getType());
        info.Add("victim", extractor.getVictim());
        info.Add("location", extractor.getLocation());
        info.Add("info", extractor.getReportInfo());
    }
    
    public List<string> getInfo()
    {
        return info;
    }
    
    public List<string> getKeyInfo(string key)
    {
        try
        {
            return info[key];
        }
        catch(KeyNotFoundException)
        {
            throw new System.ArgumentException("Key does not exist.")
        }
    }
}

public partial class Form1 : Form  
{
    InfoHandler handler = new InfoHandler();
    
    public Form1()  
    {  
        InitializeComponent();  
    }  

    private void button1_Click(object sender, EventArgs e)  
    { 
        SqlConnection con = new SqlConnection(@"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database1.mdf;Integrated Security=True;User Instance=True");  
        SqlCommand cmd = new SqlCommand("sp_insert", con);
        
        cmd.CommandType = CommandType.StoredProcedure;
        
        cmd.Parameters.AddWithValue("@CR", handler.info["cr"]);  
        cmd.Parameters.AddWithValue("@date", handler.info["date"]);  
        cmd.Parameters.AddWithValue("@officer", handler.info["officer"]);  
        cmd.Parameters.AddWithValue("@report type", handler.info["type"]); 
        cmd.Parameters.AddWithValue("@victim", handler.info["victim"]);
        cmd.Parameters.AddWithValue("@location", handler.info["location"]);
        cmd.Parameters.AddWithValue("@report", handler.info["report"]);
        
        con.Open();  
        
        int i = cmd.ExecuteNonQuery();  

        con.Close();  
 
        if (i!=0)  
        {  
            MessageBox.Show(i + "Data Saved");   
        }  

    }  

    public static void main(string[] args)  
    {  
        Application.Run(new Form1());  
    }  
}  
