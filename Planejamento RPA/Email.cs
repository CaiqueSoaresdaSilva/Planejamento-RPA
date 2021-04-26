using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Planejamento_RPA
{
    class Email
    {

        string emailDestino;

        public void EnviarEmail(string assunto, string mensagem)
        {
            //l.LogInfoEmail();
            try
            {
                DataTable DT = new DataTable();
                clDALSQL clDAL = new clDALSQL();
                MailMessage mailMessage = new MailMessage();

                //Endereço que irá aparecer no e-mail do usuário 
                mailMessage.From = new MailAddress("tecnologia@credgroup.com.br");


                clDAL.Ambiente = clDALSQL.AmbienteExecucao.Producao;
                clDAL.ClearParameters();

                DT = clDAL.RetornaTabela("dbo.PRC_DISCADOR_LISTA_EMAILS", "MANAGER");
                foreach (DataRow DR in DT.Rows)
                {
                    emailDestino = DR["EMAIL"].ToString();
                    mailMessage.To.Add(emailDestino);
                }

                mailMessage.Subject = assunto;
                mailMessage.IsBodyHtml = true;

                //conteudo do corpo do e-mail 
                mailMessage.Body = mensagem;
                mailMessage.Priority = MailPriority.Normal;

                //smtp do e-mail que irá enviar 
                SmtpClient smtpClient = new SmtpClient("smtplw.com.br");
                smtpClient.EnableSsl = false;
                smtpClient.Port = 587;

                //credenciais da conta que utilizará para enviar o e-mail 
                smtpClient.Credentials = new NetworkCredential("systemcredenvio", "xTdlkDfd7827");
                smtpClient.Send(mailMessage);
                //l.LogEmailOK();

            }
            catch (Exception ex)
            {
                ex.Message.ToString();
                throw;
            }
        }

    }

}
