using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Collections;
using com.salesforce.eu2;

public partial class test : System.Web.UI.Page
{
    private OleDbConnection con;
    private OleDbCommand cmd;
    private String query, username, password;
    SforceService SfdcBinding = null;
    LoginResult CurrentLoginResult = null;

    protected void Page_Load(object sender, EventArgs e)
    {
        query = @"SELECT functiecode FROM [Access].[100_bellijst] WHERE status = 'Verwijderd'";
        QueryMaker(query);

        OleDbDataReader sqlReader = cmd.ExecuteReader();
        while (sqlReader.Read())
        {
          GetBOData(sqlReader.GetString(0)); 
        } 
    }
 
    //Fetching Data and Creating lead
    public void GetBOData(String Functiecode)
    {
        query = @"SELECT [functietitel], [bedrijfsnaam], [locatie], [functieomschrijving], [bron], [opmerking],
                [contactpersoon_naam], [contactpersoon_mobiel], [contactpersoon_tel], [contactpersoon_e_mail],
                [intermediair], [functiecode] FROM [Access].[100_bellijst] WHERE functiecode = '" + Functiecode + "'";

        QueryMaker(query);
        OleDbDataReader sqlReader = cmd.ExecuteReader();
        while (sqlReader.Read())
            {
                String functietitel = "" + sqlReader.GetValue(0) + "";
                String bedrijfsnaam = "" + sqlReader.GetValue(1) + "";
                String locatie = "" + sqlReader.GetValue(2) + "";
                String functieomschrijving = "" + sqlReader.GetValue(3) + "";
                String bron = "" + sqlReader.GetValue(4) + "";
                String opmerking = "" + sqlReader.GetValue(5) + "";
                String contactpersoon_naam = "" + sqlReader.GetValue(6) + "";
                String contactpersoon_mobiel = "" + sqlReader.GetValue(7) + "";
                String contactpersoon_tel = "" + sqlReader.GetValue(8) + "";
                String contactpersoon_e_mail = "" + sqlReader.GetValue(9) + "";
                String intermediar = "" + sqlReader.GetValue(10) + "";
                String functiecode = "" + sqlReader.GetValue(11) + "";  

               CreateLead(functiecode, bedrijfsnaam, locatie, contactpersoon_naam, contactpersoon_mobiel, contactpersoon_tel, contactpersoon_e_mail, intermediar, bron);
            }
      con.Close();
    }

    public void CreateLead(String functiecode,String bedrijfsnaam, String locatie, String contactpersoon_naam, String contactpersoon_mobiel, String contactpersoon_tel, String contactpersoon_e_mail, String intermediar, String bron)
    {

        //SF Connectie
        SFSoapAPICon();

        /*Contactpersoon_naam = max 80
        Bedrijfsnaam of Intermediair = max 255
        Contactpersoon_email = max 80
        Locatie = max 40
        Country = max 80
        contactpersoon_tel = max 40
        contactpersoon_mobiel = max 40
        Website = max 255
        Lead source = max 40 */

        //New Lead
        Lead sfdcLead = new Lead();

        sfdcLead.FirstName = "";
        sfdcLead.LastName = contactpersoon_naam;
       // Note note = new Note();
        // sfdcLead.Notes = Query;

        if (bedrijfsnaam.Count() == 0 || intermediar.Count() == 0)
        {
            sfdcLead.Company = "Onbekend";
        }
        else if (bedrijfsnaam.Count() > 0 || intermediar.Count() > 0)
        {
            sfdcLead.Company = bedrijfsnaam;
        }
        else
        {
            if (bedrijfsnaam.Count() == 0)
            {
                sfdcLead.Company = intermediar;
                Response.Write("Intermediar" + intermediar + "<br/>");
            }
            else
            {
                sfdcLead.Company = bedrijfsnaam;
                Response.Write("bedrijfsnaam" + bedrijfsnaam + "<br/>");
            }
        }

        sfdcLead.Company_phone__c = contactpersoon_tel;
        sfdcLead.Email = contactpersoon_e_mail;
        sfdcLead.City = locatie;
        sfdcLead.Country = "Nederland";
        sfdcLead.Phone = contactpersoon_tel;
        sfdcLead.MobilePhone = contactpersoon_mobiel;
       // sfdcLead.Website = "";
        sfdcLead.LeadSource = "Freelance.nl";

        //Note
        //Functiecode, Functietitel, functieomschrijving, bron

        //Send Result to SF
        SaveResult[] saveResults = SfdcBinding.create(new sObject[] { sfdcLead });
        
        if (saveResults[0].success)
        {
            string Id = "";
            Id = saveResults[0].id;
            System.Diagnostics.Debug.WriteLine(sfdcLead);
        }
        else
        {
            string result = "";
            result = saveResults[0].errors[0].message;
            System.Diagnostics.Debug.WriteLine(result);

            query = @"UPDATE [Access].[100_bellijst] set [status] = 'Aanvullen' WHERE functiecode = '" + functiecode + "'";
            QueryMaker(query);
            cmd.ExecuteNonQuery();
        }
    }

    //Query maker
    public void QueryMaker(String query)
    {
        //Acces acces database
        con = new OleDbConnection(@"Provider=SQLNCLI11; 
        Password=29^}$4p=$7%4g1A; 
        User ID=Summaview@qf1kzf54o1;
        Initial Catalog=businessopportunities;
        Data Source =tcp:qf1kzf54o1.database.windows.net;");

        cmd = new System.Data.OleDb.OleDbCommand();
        cmd.CommandType = System.Data.CommandType.Text;
        cmd.CommandText = query;
        cmd.Connection = con;
        con.Open();
    }

    public void SFSoapAPICon()
    {
        username = "vrieling@visumma.nl";
        password = "6'@$c;q<{kQVL5swgiY22NhludJCkHMcAkA";
        SfdcBinding = new SforceService();

        try
        {
            CurrentLoginResult = SfdcBinding.login(username, password);
            System.Diagnostics.Debug.WriteLine("Connected!");
        }
        catch (System.Web.Services.Protocols.SoapException e)
        {
            // This is likley to be caused by bad username or password
            SfdcBinding = null;
            throw (e);
        }
        catch (Exception e)
        {
            // This is something else, probably communication
            SfdcBinding = null;
            throw (e);
        }

        //Change the binding to the new endpoint
        SfdcBinding.Url = CurrentLoginResult.serverUrl;

        //Create a new session header object and set the session id to that returned by the login
        SfdcBinding.SessionHeaderValue = new SessionHeader();
        SfdcBinding.SessionHeaderValue.sessionId = CurrentLoginResult.sessionId;
    }
}