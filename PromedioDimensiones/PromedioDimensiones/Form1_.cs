using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Configuration;

namespace leefoxdb
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // realiza instancias necesarias para conexion bd FOX
                // --------------------------------------------------
                OleDbConnection con = new OleDbConnection();

                // Obtiene el string de conexion de la BD de fox
                // ---------------------------------------------
                string ConexionStringFX = ConfigurationManager.AppSettings["ConexionStringFX"];
                
                // variables necesarias para el control del programa
                // -------------------------------------------------
                // variable contendra el contenido del correo
                // -------------------------------------------
                string BodyMail = "";

                // variable para controlar el fondo de las filas del reporte
                // ---------------------------------------------------------
                int contador = 1;

                // crea el data table para colocar los resultados de la consulta
                // -------------------------------------------------------------
                DataTable dt = new DataTable();

                // asigna el el string de conexion
                // --------------------------------
                con.ConnectionString = ConexionStringFX;

                // abre conexion de la base de datos
                // ---------------------------------
                con.Open();  

                // realiza instancia para enviar parametros a la BD
                // ------------------------------------------------
                OleDbCommand ocmd = con.CreateCommand();

                // arma query para obtener resultado
                // ---------------------------------
                //ocmd.CommandText = @"SELECT a.guia,c.razons,c.account,a.peso,a.bultos,a.fechai,a.fechas,DATE()-fechai dias FROM tguias a, torden b, scca02 c WHERE b.guia=a.guia AND b.tramite='DDP' AND EMPTY(fechas) AND b.codcte=c.codcte AND (DATE()-fechai)>=2";
                ocmd.CommandText = @"SELECT a.guia,c.razons,c.account,a.peso,a.bultos,a.fechai,a.fechas,DATE()-fechai dias FROM tguias a, torden b, scca02 c WHERE b.guia = a.guia AND b.tramite = 'IMP' AND EMPTY(fechas) AND b.codcte = c.codcte AND(DATE() - fechai) >= 2";
                // ejecuta y carga a data table el resultado de la consulta
                // --------------------------------------------------------
                dt.Load(ocmd.ExecuteReader());

                // arma tabla para colocar los registros encontrados de la consulta
                // ----------------------------------------------------------------
                BodyMail = @"<table> <tr bgcolor= ""#1B27A7"" style=""color:#ffffff"">";
                BodyMail = BodyMail +"<td>Guia </td> ";
                BodyMail = BodyMail + "<td>Razon </td> ";
                BodyMail = BodyMail + "<td>Cuenta </td> ";
                BodyMail = BodyMail + "<td>Peso </td> ";
                BodyMail = BodyMail + "<td>Bulto </td> ";
                BodyMail = BodyMail + "<td>Fecha Ingreso </td> ";
                BodyMail = BodyMail + "<td>Fecha Salida </td> ";
                BodyMail = BodyMail + "<td>Dias </td> ";
                BodyMail = BodyMail + "</tr>";

                // Recorro las filas obtenidas de la consulta
                // ------------------------------------------
                foreach (DataRow row in dt.Rows)
                {
                    // obtengo valores de la fila procesada
                    // ------------------------------------
                    string guia     = row["guia"].ToString();
                    string razons   = row["razons"].ToString();
                    string account  = row["account"].ToString();
                    string peso     = row["peso"].ToString();
                    string bultos   = row["bultos"].ToString();
                    string fechai   = row["fechai"].ToString();
                    string fechas   = row["fechas"].ToString();
                    string dias     = row["dias"].ToString();

                    // si viene una fecha invalida inicializa campo de fecha de salida
                    // ---------------------------------------------------------------
                    if (fechas == "30/12/1899 0:00:00")
                        fechas = "";

                    // arma la fila con el color de fondo que corresponde
                    // --------------------------------------------------
                    if (contador == 1)
                        BodyMail = BodyMail + @"<tr bgcolor= ""#B3E7FF"" >";
                    else
                        BodyMail = BodyMail + @"<tr bgcolor= ""#EAF8FF"" >";

                    // agrega valores a el cuerpo del correo
                    // -------------------------------------
                    BodyMail = BodyMail + "<td> ";
                    BodyMail = BodyMail + guia;
                    BodyMail = BodyMail + "</td> ";
                    BodyMail = BodyMail + "<td> ";
                    BodyMail = BodyMail + razons;
                    BodyMail = BodyMail + "</td> ";
                    BodyMail = BodyMail + "<td> ";
                    BodyMail = BodyMail + account;
                    BodyMail = BodyMail + "</td> ";
                    BodyMail = BodyMail + "<td> ";
                    BodyMail = BodyMail + peso;
                    BodyMail = BodyMail + "</td> ";
                    BodyMail = BodyMail + "<td> ";
                    BodyMail = BodyMail + bultos;
                    BodyMail = BodyMail + "</td> ";
                    BodyMail = BodyMail + "<td> ";
                    BodyMail = BodyMail + fechai;
                    BodyMail = BodyMail + "</td> ";
                    BodyMail = BodyMail + "<td> ";
                    BodyMail = BodyMail + fechas;
                    BodyMail = BodyMail + "</td> ";
                    BodyMail = BodyMail + "<td> ";
                    BodyMail = BodyMail + dias;
                    BodyMail = BodyMail + "</td> ";
                    BodyMail = BodyMail + "</tr>";

                    // incrementa contador para saber el color de linea que corresponde a la fila procesada
                    // ------------------------------------------------------------------------------------
                    contador = contador + 1;

                    // solo se tienen dos colores por lo que si sobrepasa de 2 inicializa el contador
                    // ------------------------------------------------------------------------------
                    if (contador > 2)
                        contador = 1;
                }

                // cierra fila y tabla para finalizar contenido
                // --------------------------------------------
                BodyMail = BodyMail + "</tr>";
                BodyMail = BodyMail + "</table>";

                // cierra conexion a la BD de FOX
                // ------------------------------
                con.Close();

                // realiza envio de correo con la informacion recopilada
                // -----------------------------------------------------
                // obtiene configuracion del config para conectarse al servidor de correo y obtener correos origen y destino
                // ---------------------------------------------------------------------------------------------------------
                string ServidorExchange = ConfigurationManager.AppSettings["ServidorExchange"];
                string UsuarioExchange = ConfigurationManager.AppSettings["UsuarioExchange"];
                string PassExchange = ConfigurationManager.AppSettings["PassExchange"];
                string CorreoOrigen = ConfigurationManager.AppSettings["CorreoOrigen"];
                string CorreoDestino = ConfigurationManager.AppSettings["CorreoDestino"];
                string NomCorreoOrigen = ConfigurationManager.AppSettings["NomCorreoOrigen"];
                string NomCorreoDestino = ConfigurationManager.AppSettings["NomCorreoDestino"];
                string AsuntoCorreo = ConfigurationManager.AppSettings["AsuntoCorreo"];
                string PuertoSMS = ConfigurationManager.AppSettings["PuertoSMS"];
                int PuertoCastSMS = Int32.Parse(PuertoSMS);
                SmtpClient mySmtpClient = null;

                if (PuertoSMS != "N/A")
                {
                    // realiza conexion al servidor de mails
                    // -------------------------------------
                    mySmtpClient = new SmtpClient(ServidorExchange, PuertoCastSMS);
                }
                else
                {
                    // realiza conexion al servidor de mails
                    // -------------------------------------
                    mySmtpClient = new SmtpClient(ServidorExchange);
                }

                // coloca las credenciales de conexion al servidor
                mySmtpClient.UseDefaultCredentials = false;
                System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential(UsuarioExchange, PassExchange);
                mySmtpClient.Credentials = basicAuthenticationInfo;

                // agrega direccion de correo origen y destino
                // -------------------------------------------
                MailAddress from = new MailAddress(CorreoOrigen, NomCorreoOrigen);
                MailAddress to = new MailAddress(CorreoDestino, NomCorreoDestino);
                MailMessage myMail = new System.Net.Mail.MailMessage(from, to);

                // agrega replyto
                // --------------
                MailAddress replyto = new MailAddress("reply@example.com");
                myMail.ReplyToList.Add(replyto);

                // agrega asunto del correo
                // ------------------------
                myMail.Subject = AsuntoCorreo;
                myMail.SubjectEncoding = System.Text.Encoding.UTF8;

                // agrega cuerpo del correo
                // ------------------------
                myMail.Body = BodyMail;
                myMail.BodyEncoding = System.Text.Encoding.UTF8;
                
                // indica si el cuerpo es texto o html
                // -----------------------------------
                myMail.IsBodyHtml = true;

                // envia correo
                // ------------
                mySmtpClient.Send(myMail);

            }
            catch (SystemException exp) 
            {
                MessageBox.Show("Error: " + exp.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string ServidorExchange = ConfigurationManager.AppSettings["ServidorExchange"];
                string UsuarioExchange  = ConfigurationManager.AppSettings["UsuarioExchange"];
                string PassExchange     = ConfigurationManager.AppSettings["PassExchange"];
                string CorreoOrigen     = ConfigurationManager.AppSettings["CorreoOrigen"];
                string CorreoDestino    = ConfigurationManager.AppSettings["CorreoDestino"];
                string NomCorreoOrigen  = ConfigurationManager.AppSettings["NomCorreoOrigen"];
                string NomCorreoDestino = ConfigurationManager.AppSettings["NomCorreoDestino"];

                SmtpClient mySmtpClient = new SmtpClient(ServidorExchange);

                // set smtp-client with basicAuthentication
                mySmtpClient.UseDefaultCredentials = false;
                System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential(UsuarioExchange, PassExchange);
                mySmtpClient.Credentials = basicAuthenticationInfo;

                // add from,to mailaddresses
                MailAddress from = new MailAddress(CorreoOrigen, NomCorreoOrigen);
                MailAddress to = new MailAddress(CorreoDestino, NomCorreoDestino);
                MailMessage myMail = new System.Net.Mail.MailMessage(from, to);

                // add ReplyTo
                MailAddress replyto = new MailAddress("reply@example.com");
                myMail.ReplyToList.Add(replyto);

                // set subject and encoding
                myMail.Subject = "Test message";
                myMail.SubjectEncoding = System.Text.Encoding.UTF8;

                // set body-message and encoding
                myMail.Body = "<b>Test Mail</b><br>using <b>HTML</b>.";
                myMail.BodyEncoding = System.Text.Encoding.UTF8;
                // text or html
                myMail.IsBodyHtml = true;

                mySmtpClient.Send(myMail);
            }
            catch (ExecutionEngineException) { }
        }
    }
}
