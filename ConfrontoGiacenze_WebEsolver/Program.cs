using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConfrontoGiacenze_WebEsolver
{
    static class Program
    {
        /// <summary>
        /// Punto di ingresso principale dell'applicazione.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            if (args[0] == "B")
            {
                string pathINI = Directory.GetCurrentDirectory() + "\\ConfrontoGiacenzeINI.ini";
                INIFile fileini = new INIFile(pathINI);
                int[] giacEsolver = new int[30];
                SqlConnection connWeb;
                int anomalia = 0;


                if (File.Exists(Directory.GetCurrentDirectory() + "\\ErrorLOG.log"))
                {
                    File.Delete(Directory.GetCurrentDirectory() + "\\ErrorLOG.log");
                }

                if (File.Exists(Directory.GetCurrentDirectory() + "\\LOGFILE.log"))
                {
                    File.Delete(Directory.GetCurrentDirectory() + "\\LOGFILE.log");
                }

                //RIGA LOG PROCEDURA
                using (StreamWriter w = File.AppendText(Directory.GetCurrentDirectory() + "\\LOGFILE.log"))
                {
                    w.WriteLine(" Inzio processo");
                    w.Flush();
                    w.Close();
                }

                //RIGA LOG PROCEDURA
                using (StreamWriter w = File.AppendText(Directory.GetCurrentDirectory() + "\\LOGFILE.log"))
                {
                    w.WriteLine("Connessione DB in corso...");
                    w.Flush();
                    w.Close();
                }

                try
                {
                    String strconnWeb = "Server = '" + fileini.IniReadValue("dsweb", "Sezione1") + "'; Database ='" + fileini.IniReadValue("catweb", "Sezione1") + "'; User Id = '" + fileini.IniReadValue("userweb", "Sezione1") + "'; Password = '" + fileini.IniReadValue("pwdweb", "Sezione1") + "';";
                    connWeb = new SqlConnection(strconnWeb);
                    connWeb.Open();

                    //RIGA LOG PROCEDURA
                    using (StreamWriter w = File.AppendText(Directory.GetCurrentDirectory() + "\\LOGFILE.log"))
                    {
                        w.WriteLine("Connessione effettuata");
                        w.Flush();
                        w.Close();
                    }

                    //RIGA LOG PROCEDURA
                    using (StreamWriter w = File.AppendText(Directory.GetCurrentDirectory() + "\\LOGFILE.log"))
                    {
                        w.WriteLine("Inizio controllo giacenze Web / Esolver...");
                        w.Flush();
                        w.Close();
                    }

                    //RECORDSET GIACENZE WEB
                    String sqlWeb = @"SELECT CodiceArticolo, Varianti.Codice, CodiceMagazzino, Quantita1, Quantita2, Quantita3, Quantita4, Quantita5, Quantita6, Quantita7, Quantita8, Quantita9, Quantita10,
Quantita11, Quantita12, Quantita13, Quantita14, Quantita15, Quantita16, Quantita17, Quantita18, Quantita19, Quantita20,
Quantita21, Quantita22, Quantita23, Quantita24, Quantita25, Quantita26, Quantita27, Quantita28, Quantita29, Quantita30,
Taglia1, Taglia2, Taglia3, Taglia4, Taglia5, Taglia6, Taglia7, Taglia8, Taglia9, Taglia10,
Taglia11, Taglia12, Taglia13, Taglia14, Taglia15, Taglia16, Taglia17, Taglia18, Taglia19, Taglia20,
Taglia21, Taglia22, Taglia23, Taglia24, Taglia25, Taglia26, Taglia27, Taglia28, Taglia29, Taglia30
FROM Varianti 
INNER JOIN Articoli on Varianti.CodiceArticolo = articoli.Codice
INNER JOIN Segnataglie ON Articoli.CodiceSegnataglie = Segnataglie.Codice
WHERE CodiceMagazzino <> '0' 
ORDER BY CodiceArticolo, Codice, CodiceMagazzino  ";
                    SqlCommand command = new SqlCommand(sqlWeb, connWeb);
                    command.CommandType = CommandType.Text;

                    SqlDataReader myReader = command.ExecuteReader();

                    if (myReader.HasRows)
                    {
                        while (myReader.Read())
                        {
                            //se ci sono record nella tabella web Varianti, creo il recordset con i dati di Esolver
                            String query1 = @"SELECT 
                                sum(ModaProgMagCor.GtgQgp1) As [GtgQgp1],  
                                sum(ModaProgMagCor.GtgQgp2) As[GtgQgp2],
                                sum(ModaProgMagCor.GtgQgp3) As [GtgQgp3],  
                                sum(ModaProgMagCor.GtgQgp4) As [GtgQgp4],  
                                sum(ModaProgMagCor.GtgQgp5) As [GtgQgp5],  
                                sum(ModaProgMagCor.GtgQgp6) As [GtgQgp6],
                                sum(ModaProgMagCor.GtgQgp7) As [GtgQgp7],
                                sum(ModaProgMagCor.GtgQgp8) As [GtgQgp8],  
                                sum(ModaProgMagCor.GtgQgp9) As [GtgQgp9], 
                                sum(ModaProgMagCor.GtgQgp10) As [GtgQgp10],
                                sum(ModaProgMagCor.GtgQgp11) As [GtgQgp11],  
                                sum(ModaProgMagCor.GtgQgp12) As [GtgQgp12],  
                                sum(ModaProgMagCor.GtgQgp13) As [GtgQgp13],  
                                sum(ModaProgMagCor.GtgQgp14) As [GtgQgp14],  
                                sum(ModaProgMagCor.GtgQgp15) As [GtgQgp15],  
                                sum(ModaProgMagCor.GtgQgp16) As [GtgQgp16],  
                                sum(ModaProgMagCor.GtgQgp17) As [GtgQgp17],  
                                sum(ModaProgMagCor.GtgQgp18) As [GtgQgp18],  
                                sum(ModaProgMagCor.GtgQgp19) As [GtgQgp19],  
                                sum(ModaProgMagCor.GtgQgp20) As [GtgQgp20],  
                                sum(ModaProgMagCor.GtgQgp21) As [GtgQgp21],  
                                sum(ModaProgMagCor.GtgQgp22) As [GtgQgp22],  
                                sum(ModaProgMagCor.GtgQgp23) As [GtgQgp23],  
                                sum(ModaProgMagCor.GtgQgp24) As [GtgQgp24],  
                                sum(ModaProgMagCor.GtgQgp25) As [GtgQgp25],  
                                sum(ModaProgMagCor.GtgQgp26) As [GtgQgp26],  
                                sum(ModaProgMagCor.GtgQgp27) As [GtgQgp27],  
                                sum(ModaProgMagCor.GtgQgp28) As [GtgQgp28],  
                                sum(ModaProgMagCor.GtgQgp29) As [GtgQgp29],  
                                sum(ModaProgMagCor.GtgQgp30) As [GtgQgp30]  
                            FROM ModaProgMagCor As [ModaProgMagCor]
                            WHERE ModaProgMagCor.GtgCart = @ARTICOLO 
                                AND ModaProgMagCor.GtgVarart = @VARIANTE
                                AND ModaProgMagCor.Gtgcmag = @MAGAZZINO
                                AND ModaProgMagCor.DBGruppo= 'CC' ";

                            String SQL = "SELECT ";
                            SQL = SQL + " SUM(QtaTagDaSpedire_1) AS QtaTagDaSpedire_1, SUM(QtaTagDaSpedire_2) AS QtaTagDaSpedire_2, SUM(QtaTagDaSpedire_3) AS QtaTagDaSpedire_3, SUM(QtaTagDaSpedire_4) AS QtaTagDaSpedire_4, SUM(QtaTagDaSpedire_5) AS QtaTagDaSpedire_5, SUM(QtaTagDaSpedire_6) AS QtaTagDaSpedire_6, SUM(QtaTagDaSpedire_7) AS QtaTagDaSpedire_7, SUM(QtaTagDaSpedire_8) AS QtaTagDaSpedire_8, SUM(QtaTagDaSpedire_9) AS QtaTagDaSpedire_9, SUM(QtaTagDaSpedire_10) AS QtaTagDaSpedire_10, ";
                            SQL = SQL + " SUM(QtaTagDaSpedire_11) AS QtaTagDaSpedire_11, SUM(QtaTagDaSpedire_12) AS QtaTagDaSpedire_12, SUM(QtaTagDaSpedire_13) AS QtaTagDaSpedire_13, SUM(QtaTagDaSpedire_14) AS QtaTagDaSpedire_14, SUM(QtaTagDaSpedire_15) AS QtaTagDaSpedire_15, SUM(QtaTagDaSpedire_16) AS QtaTagDaSpedire_16, SUM(QtaTagDaSpedire_17) AS QtaTagDaSpedire_17, SUM(QtaTagDaSpedire_18) AS QtaTagDaSpedire_18, SUM(QtaTagDaSpedire_19) AS QtaTagDaSpedire_19, SUM(QtaTagDaSpedire_20) AS QtaTagDaSpedire_20, ";
                            SQL = SQL + " SUM(QtaTagDaSpedire_21) AS QtaTagDaSpedire_21, SUM(QtaTagDaSpedire_22) AS QtaTagDaSpedire_22, SUM(QtaTagDaSpedire_23) AS QtaTagDaSpedire_23, SUM(QtaTagDaSpedire_24) AS QtaTagDaSpedire_24, SUM(QtaTagDaSpedire_25) AS QtaTagDaSpedire_25, SUM(QtaTagDaSpedire_26) AS QtaTagDaSpedire_26, SUM(QtaTagDaSpedire_27) AS QtaTagDaSpedire_27, SUM(QtaTagDaSpedire_28) AS QtaTagDaSpedire_28, SUM(QtaTagDaSpedire_29) AS QtaTagDaSpedire_29, SUM(QtaTagDaSpedire_30) AS QtaTagDaSpedire_30 ";
                            SQL = SQL + " FROM OrdSpedTestata ";
                            SQL = SQL + " INNER JOIN OrdSpedRighe On OrdSpedTestata.IdDocumento = OrdSpedRighe.IdDocumento ";
                            SQL = SQL + " INNER JOIN ModaSpedRighe On ModaSpedRighe.IddOrdSped = OrdSpedRighe.IdDocumento And ModaSpedRighe.IdrOrdSped = OrdSpedRighe.IdRiga ";
                            SQL = SQL + " WHERE OrdSpedTestata.StatoSpedizione < 4 And CodArt = '" + myReader["CodiceArticolo"].ToString() + "' AND VarianteArt = '" + myReader["Codice"].ToString() + "' AND CodMag = '" + myReader["CodiceMagazzino"].ToString() + "'";

                            using (SqlConnection esolver = new SqlConnection("Server = '" + fileini.IniReadValue("dstsm", "Sezione1") + "'; Database ='" + fileini.IniReadValue("cattsm", "Sezione1") + "'; User Id = '" + fileini.IniReadValue("usertsm", "Sezione1") + "'; Password = '" + fileini.IniReadValue("pwdtsm", "Sezione1") + "';"))
                            {
                                esolver.Open();

                                SqlCommand giacMag = new SqlCommand(query1, esolver);
                                giacMag.CommandType = CommandType.Text;

                                giacMag.Parameters.Add("@ARTICOLO", SqlDbType.VarChar);
                                giacMag.Parameters["@ARTICOLO"].Value = myReader["CodiceArticolo"].ToString();

                                giacMag.Parameters.Add("@VARIANTE", SqlDbType.VarChar);
                                giacMag.Parameters["@VARIANTE"].Value = myReader["Codice"].ToString();

                                giacMag.Parameters.Add("@MAGAZZINO", SqlDbType.VarChar);
                                giacMag.Parameters["@MAGAZZINO"].Value = myReader["CodiceMagazzino"].ToString();

                                SqlDataReader GiacNetta = giacMag.ExecuteReader();

                                if (GiacNetta.HasRows)
                                {
                                    GiacNetta.Read();

                                    for (int i = 1; i <= 30; i++)
                                    {
                                        if (GiacNetta["GtgQgp" + i.ToString()] != DBNull.Value)
                                        {
                                            if (Convert.ToInt32(GiacNetta["GtgQgp" + i.ToString()]) > 0)
                                            {
                                                giacEsolver[i - 1] = Convert.ToInt32(GiacNetta["GtgQgp" + i.ToString()]);
                                            }
                                            else { giacEsolver[i - 1] = 0; }
                                        }
                                        else
                                        { giacEsolver[i - 1] = 0; }
                                    }
                                }

                                GiacNetta.Close();
                            }

                            //            MsgAddLog("finito il primo inizio con il secondo");

                            using (SqlConnection esolver = new SqlConnection("Server = '" + fileini.IniReadValue("dstsm", "Sezione1") + "'; Database ='" + fileini.IniReadValue("cattsm", "Sezione1") + "'; User Id = '" + fileini.IniReadValue("usertsm", "Sezione1") + "'; Password = '" + fileini.IniReadValue("pwdtsm", "Sezione1") + "';"))
                            {
                                esolver.Open();

                                SqlCommand giacMag = new SqlCommand(SQL, esolver);
                                giacMag.CommandType = CommandType.Text;

                                SqlDataReader Prelevato = giacMag.ExecuteReader();

                                if (Prelevato.HasRows)
                                {
                                    Prelevato.Read();

                                    for (int i = 1; i <= 30; i++)
                                    {
                                        if (Prelevato["QtaTagDaSpedire_" + i.ToString()] != DBNull.Value)
                                        {
                                            if (Convert.ToInt32(Prelevato["QtaTagDaSpedire_" + i.ToString()]) > 0)
                                            {
                                                giacEsolver[i - 1] = giacEsolver[i - 1] - Convert.ToInt32(Prelevato["QtaTagDaSpedire_" + i.ToString()]);

                                                if (giacEsolver[i - 1] < 0)
                                                {
                                                    giacEsolver[i - 1] = 0;
                                                }
                                            }
                                        }
                                    }
                                }

                                Prelevato.Close();
                            }

                            //controllo se quantita web <> quantita esolver taglia x taglia 
                            for (int i = 1; i <= 30; i++)
                            {
                                if (Convert.ToInt32(myReader["Quantita" + i.ToString()]) != giacEsolver[i - 1])
                                {
                                    //Quantita diverse = scrivo su file articolo, variante e magazzino
                                    using (StreamWriter w = File.AppendText(Directory.GetCurrentDirectory() + "\\ErrorLOG.log"))
                                    {
                                            w.Write("\r\n ");
                                            w.Write("  :{0} / {1} / {2} / Taglia {3} / Qta web: {4} / Qta es: {5}", myReader["CodiceArticolo"].ToString(), myReader["Codice"].ToString(), myReader["CodiceMagazzino"].ToString(), myReader["Taglia" + i.ToString()].ToString(), myReader["Quantita" + i.ToString()].ToString(), giacEsolver[i - 1].ToString());
                                            w.Flush();
                                            w.Close();
                                    }

                                    anomalia = 1;
                                }
                            }
                        }
                    }

                    myReader.Close();

                    //RIGA LOG PROCEDURA
                    using (StreamWriter w = File.AppendText(Directory.GetCurrentDirectory() + "\\LOGFILE.log"))
                    {
                        w.WriteLine("Invio mail in corso...");
                        w.Flush();
                        w.Close();
                    }

                    try
                    {
                        MailMessage myMail = new MailMessage();
                        myMail.From = new MailAddress("mattia.datasoft@gmail.com");
                        myMail.To.Add("m.donini@datasoft.it, DavideCittone@add.it");
                        myMail.CC.Add("v.righi@datasoft.it");
                        myMail.Subject = "Controllo giacenze Esolver / Web -- " + fileini.IniReadValue("Web", "Sezione1");
                        myMail.Priority = MailPriority.Normal;

                        if (anomalia == 1)
                        {
                            myMail.Body = "Controlla file di log in allegato per le segnalazioni!";
                            myMail.Attachments.Add(new Attachment(Directory.GetCurrentDirectory() + "\\ErrorLOG.log"));
                        }
                        else
                        {
                            myMail.Body = "Controllo andato a buon fine!";
                        }

                        SmtpClient Smtp = new SmtpClient("smtp.gmail.com");
                        Smtp.Port = 587;
                        Smtp.EnableSsl = true;
                        Smtp.Credentials = new NetworkCredential("mattia.datasoft@gmail.com", "DataSoft2017");
                        Smtp.Send(myMail);
                    }
                    catch (Exception ex)
                    {
                        //RIGA LOG PROCEDURA
                        using (StreamWriter w = File.AppendText(Directory.GetCurrentDirectory() + "\\LOGFILE.log"))
                        {
                            w.WriteLine("Errore nell'invio della mail");
                            w.Flush();
                            w.Close();
                        }
                    }

                    //RIGA LOG PROCEDURA
                    using (StreamWriter w = File.AppendText(Directory.GetCurrentDirectory() + "\\LOGFILE.log"))
                    {
                        w.WriteLine("Fine controllo giacenze Web / Esolver");
                        w.Flush();
                        w.Close();
                    }



                }
                catch (Exception ex)
                {
                    //RIGA LOG PROCEDURA
                    using (StreamWriter w = File.AppendText(Directory.GetCurrentDirectory() + "\\LOGFILE.log"))
                    {
                        w.WriteLine("ERRORE PROCEDURA: -- " + ex.Message);
                        w.Flush();
                        w.Close();
                    }
                }


            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
        }
    }
}
