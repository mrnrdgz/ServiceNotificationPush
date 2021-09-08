using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Security;
using System.Security.Permissions;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using Google.Apis.Auth.OAuth2;
using System.Net.Http;
using NLog;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        //private static string connectionString = @"Data Source=GEA_PICO;Initial Catalog=GeaCorpico;User ID=sa;Password=bmast24";
        private static string connectionString = @"Data Source=.;Initial Catalog=GeaCorpico;Integrated Security = True";
        private static Dictionary<string, SqlDependency> sqlDependencies = new Dictionary<string, SqlDependency>();
        private static SqlConnection oConnection = null;
        private static int ultimaNotificacion = DateTime.Now.Millisecond;
        private static string ultimoTopic = null;
        private static List<string> topicList = new List<string>();
        private static NLog.Logger logger  = NLog.LogManager.GetCurrentClassLogger();

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                DemandSqlPermission();
                StartSqlServerObservation();            
                RegisterForChanges();
                observeNetworkStatus();
            }
            catch (SecurityException e)
            {
                EventLog.WriteEntry(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss.fff") + " No tiene persmisos SQLClient...." + e.Message);
                logger.Error(e.Message, "No tiene persmisos SQLClient....");
            }
            catch (SqlException exception)
            {
                EventLog.WriteEntry(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss.fff") + " - EXCEPCION SQL: " + exception.Message);
                logger.Error(exception.Message, "EXCEPCION SQL: ");
                sendLogMail();
                reLoadLog();
                RegisterForChanges();

            }
            catch (InvalidOperationException e)
            {
                EventLog.WriteEntry(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss.fff") + " - EXCEPCION SQL: " + e.Message);
                logger.Error(e.Message, "EXCEPCION SQL: ");
                sendLogMail();
                reLoadLog();
                RegisterForChanges();
            }
        }
        private void DemandSqlPermission()
        {
            SqlClientPermission permission = new SqlClientPermission(PermissionState.Unrestricted);
            permission.Demand();
            EventLog.WriteEntry(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss.fff") + " DemandSqlPermission....");
        }
        private void StartSqlServerObservation()
        {
            bool sqlServerOff = true;

            while (sqlServerOff)
            {
                try
                {
                    SqlDependency.Start(connectionString);
                    sqlServerOff = false;
                    EventLog.WriteEntry(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss.fff") + " - StartSqlServerObservation ");
                }
                catch (SqlException innerE)
                {
                    EventLog.WriteEntry(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss.fff") + " - Error al conectar con SQL SERVER - Próximo reintento en 10s..." + innerE.Message);
                    logger.Error(innerE.Message, "Error al conectar con SQL SERVER - Próximo reintento en 10s...");
                    Thread.Sleep(10000);
                }
            }
        }
        public static void RegisterForChanges()
        {
            // Connecting to the database using our subscriber connection string 
            // and waiting for changes...
            oConnection = new SqlConnection(connectionString);
            oConnection.Open();

            SqlCommand oCommand = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =1 OR N.OTN_EMPRESA_CONTRATISTA IN (11,12,13,14,15,18,25))", oConnection);
            SqlCommand oCommand2 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =2 OR N.OTN_EMPRESA_CONTRATISTA = 16)", oConnection);
            /*SqlCommand oCommand3 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                            [OTN_MOTIVO_TRABAJO],
                                            [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                            FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                            INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                            N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =3 OR N.OTN_EMPRESA_CONTRATISTA = 21)", oConnection);
            SqlCommand oCommand4 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                            [OTN_MOTIVO_TRABAJO],
                                            [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                            FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                            INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                            N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =4 OR N.OTN_EMPRESA_CONTRATISTA = 22)", oConnection);*/
            SqlCommand oCommand5 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =5 OR N.OTN_EMPRESA_CONTRATISTA IN (19,21,22,23,26,28,29,30,31,32))", oConnection);
            SqlCommand oCommand6 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                            [OTN_MOTIVO_TRABAJO],
                                            [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                            FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                            INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                            N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =6 OR N.OTN_EMPRESA_CONTRATISTA = 56)", oConnection);
            SqlCommand oCommand7 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                            [OTN_MOTIVO_TRABAJO],
                                            [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                            FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                            INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                            N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =7 OR  N.OTN_EMPRESA_CONTRATISTA IN (1,57))", oConnection);
            SqlCommand oCommand8 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR = 8 OR N.OTN_EMPRESA_CONTRATISTA IN (17,20))", oConnection);

            SqlCommand oCommand9 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =9 OR N.OTN_EMPRESA_CONTRATISTA IN (33,50))", oConnection);
            SqlCommand oCommand10 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =10 OR N.OTN_EMPRESA_CONTRATISTA = 34)", oConnection);
            SqlCommand oCommand11 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =11 OR N.OTN_EMPRESA_CONTRATISTA IN (35,36,38,39,40,49,51,52,55))", oConnection);
            SqlCommand oCommand12 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO], 
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =12 OR N.OTN_EMPRESA_CONTRATISTA IN (41,42,43,44,45,46,53,54))", oConnection);

            SqlCommand oCommand13 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =13 OR N.OTN_EMPRESA_CONTRATISTA = 27)", oConnection);
            SqlCommand oCommand14 = new SqlCommand(@"SELECT [OTN_ID],[OTN_ORDEN],[OTN_FECHA],[OTN_TIPO_TRABAJO],
                                                [OTN_MOTIVO_TRABAJO],
                                                [OTN_ESTADO],[OTN_EMPRESA_CONTRATISTA] 
                                                FROM [dbo].[ORDEN_OPERATIVA_NOTIFICACION] N
                                                INNER JOIN [dbo].[MOTIVO_TRABAJO_SECTOR] S ON N.OTN_TIPO_TRABAJO = S.MTS_ID_TRABAJO AND 
                                                N.OTN_MOTIVO_TRABAJO = S.MTS_ID_MOTIVO AND (S.MTS_ID_SECTOR =14 OR N.OTN_EMPRESA_CONTRATISTA = 7)", oConnection);

            sqlDependencies["1"] = new SqlDependency(oCommand);
            sqlDependencies["2"] = new SqlDependency(oCommand2);
            /*sqlDependencies["3"] = new SqlDependency(oCommand3);
            sqlDependencies["4"] = new SqlDependency(oCommand4);*/
            sqlDependencies["5"] = new SqlDependency(oCommand5);
            sqlDependencies["6"] = new SqlDependency(oCommand6);
            sqlDependencies["7"] = new SqlDependency(oCommand7);
            sqlDependencies["8"] = new SqlDependency(oCommand8);
            sqlDependencies["9"] = new SqlDependency(oCommand9);
            sqlDependencies["10"] = new SqlDependency(oCommand10);
            sqlDependencies["11"] = new SqlDependency(oCommand11);
            sqlDependencies["12"] = new SqlDependency(oCommand12);
            sqlDependencies["13"] = new SqlDependency(oCommand13);
            sqlDependencies["14"] = new SqlDependency(oCommand14);

            // Asociar escucha para cambios de base de datos
            foreach (KeyValuePair<string, SqlDependency> dep in sqlDependencies)
            {
                dep.Value.OnChange += OnNotificationChange;
            }

            // Ejecutar comandos
            DataTable table = new DataTable();
            DataTable table2 = new DataTable();
            DataTable table3 = new DataTable();
            DataTable table4 = new DataTable();
            DataTable table5 = new DataTable();
            DataTable table6 = new DataTable();
            DataTable table7 = new DataTable();
            DataTable table8 = new DataTable();
            DataTable table9 = new DataTable();
            DataTable table10 = new DataTable();
            DataTable table11 = new DataTable();
            DataTable table12 = new DataTable();
            DataTable table13 = new DataTable();
            DataTable table14 = new DataTable();

            table.Load(oCommand.ExecuteReader());
            table2.Load(oCommand2.ExecuteReader());
            /*table3.Load(oCommand3.ExecuteReader());
            table4.Load(oCommand4.ExecuteReader());*/
            table5.Load(oCommand5.ExecuteReader());
            table6.Load(oCommand6.ExecuteReader());
            table7.Load(oCommand7.ExecuteReader());
            table8.Load(oCommand8.ExecuteReader());
            table9.Load(oCommand9.ExecuteReader());
            table10.Load(oCommand10.ExecuteReader());
            table11.Load(oCommand11.ExecuteReader());
            table12.Load(oCommand12.ExecuteReader());
            table13.Load(oCommand13.ExecuteReader());
            table14.Load(oCommand14.ExecuteReader());

            //oConnection.Close();       

        }
        public static void OnNotificationChange(object caller, SqlNotificationEventArgs e)
        {

            bool changeType = SqlNotificationInfo.Insert == e.Info;

            if (e.Type == SqlNotificationType.Change)
            {

                SqlDependency dependency = caller as SqlDependency;
                // Limpiar referencia          
                dependency.OnChange -= OnNotificationChange;


                string topic = ExtractTopic(dependency);

                if (!NetworkStatus.IsAvailable)
                {
                    logger.Error("Sin Internet. No es posible enviar la Notificacion al Sector: " + topic);
                    topicList.Add(topic);
                    return;
                }
                
                logger.Info("Cambio detectado. Enviando notificación al Sector...." + topic);

                sendNotification(topic);
                ultimaNotificacion = DateTime.Now.Millisecond;
                ultimoTopic = topic;
                RegisterForChanges();

            }

        }

        private static string ExtractTopic(SqlDependency dependency)
        {
            foreach (KeyValuePair<string, SqlDependency> kvp in sqlDependencies)
            {
                if (dependency == kvp.Value)
                {
                    return kvp.Key;
                }
            }

            throw new ArgumentException();
        }
        public static async Task sendNotification(string inTopic)
        {

            // Crear petición con URL base
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, "https://fcm.googleapis.com/v1/projects/app-corpico-operario/messages:send");

            // Obtener token de registro a Firebase Cloud Messaging
            string firebaseKey = await getToken();

            // Cabecera de autorización
            request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + firebaseKey);

            // Cuerpo de la petición
            var jsonObject = new
            {
                message = new
                {
                    topic = inTopic,
                    data = new
                    {
                        body = "Actualizando datos..",
                        title = "Sync"
                    },
                    android = new
                    {
                        priority = "high"
                    },
                    webpush = new
                    {
                        headers = new
                        {
                            Urgency = "high"
                        },
                        notification = new
                        {
                            body = "This is a message from FCM to web",
                            requireInteraction = "true"

                        }
                    }
                }
            };

            // Serialización JSON a STRING
            var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(jsonObject);

            // Cabecera contenido 
            request.Content = new StringContent(jsonString, Encoding.UTF8, "application/json");

            // Enviar petición y recibir respuesta
            HttpResponseMessage result;
            //Ver el try y el catch...
            try
            {
                using (var client = new HttpClient())
                {

                    result = await client.SendAsync(request);
                    var resultInString = await result.Content.ReadAsStringAsync();

                    if (result.IsSuccessStatusCode)
                    {
                        logger.Info("Notificación Push enviada > al sector: " + inTopic);
                    }
                    else
                    {
                        logger.Error("Notificación Push enviada > error : " + resultInString);
                        sendLogMail();
                        reLoadLog();
                        RegisterForChanges();
                    }

                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, "Notificación Push NO enviada > error : ");
                sendLogMail();
                reLoadLog();
                RegisterForChanges();
            }
        }

        public static async Task<string> getToken()
        {
            GoogleCredential credential;
            var res = "";

            using (var stream = new System.IO.FileStream(@"C:\ServiceNotification2\WindowsServiceNotification\firebase-file.json",
                System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(
                    new string[]{
                        "https://www.googleapis.com/auth/firebase.messaging",
                        "https://www.googleapis.com/auth/userinfo.email"
                    });
            }
            ITokenAccess c = credential as ITokenAccess;

            try
            {
                res = await c.GetAccessTokenForRequestAsync();
            }
            catch (Exception ex)
            {                
                logger.Error(ex.Message, "Error al obtener el Token....");
                res = ex.Message;
                return res;
            }
            return res;
        }
        private void observeNetworkStatus()
        {
            NetworkStatus.AvailabilityChanged += new NetworkStatusChangedHandler(DoAvailabilityChanged);
        }
        public void DoAvailabilityChanged(object sender, NetworkStatusChangedArgs e)
        {
            ReportAvailability();
        }
        private async Task ReportAvailability()
        {

            //List<String> pendigNotificationList = new List<String>(topicList.Distinct());

            if (NetworkStatus.IsAvailable)
            {
                logger.Info("La conexión a Internet ha sido Restablecida....");
                //Mando una Notificación por aca sector que ha perdido
                await sendPendingNotification(topicList.Distinct().ToList());
                //Mando mail con el log adjunto
                sendLogMail();
                //logger = NLog.LogManager.GetCurrentClassLogger();
                reLoadLog();
                RegisterForChanges();
                
            }
        }
        public static void reLoadLog(){
           /* // save all log entries
            LogManager.Flush();
            // get configuration from NLog.config
            LogManager.Configuration = LogManager.Configuration.Reload();
            // update loggers
            LogManager.ReconfigExistingLoggers();
            logger = NLog.LogManager.GetCurrentClassLogger();*/

            LogManager.Configuration = LogManager.Configuration.Reload();
            LogManager.ReconfigExistingLoggers();
        }
        private async Task sendPendingNotification(List<String> pendingNotificationList)
        {
            foreach (String topic in pendingNotificationList)
            {
                await sendNotification(topic);
                //topicList.RemoveAll(topicLst => topicLst.Equals(topic));
            }
            topicList.Clear();
        }
        public static void sendLogMail()
        {
            List<string> lstArchivoAdjunto = new List<string>();
            String ruta = @"C:\ServiceNotificationPushLog\NotificacionesPush-";
            String fecha = DateTime.Today.ToString("yyyy-MM-dd");
            String ext = ".log";
            String archivo = ruta + fecha + ext;
            lstArchivoAdjunto.Add(archivo);

            //Todo: comprobar si existe el archivo...para hacer esto...
            //creamos nuestro objeto de la clase que hicimos
            Mail oMail = new Mail("notificacionespush.corpico@gmail.com", "mrnrdgz@gmail.com",
                                "Internet Corpico Restaurada - Notificaciones Push CorpicoApp", "Log Adjunto", lstArchivoAdjunto);

            //y enviamos
            if (oMail.enviaMail())
            {
                logger.Info("Mail enviado con Log adjunto " + archivo);
            }
            else
            {
                logger.Error(DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff") + " No se envio el mail: " + oMail.error);
            }
        }
        protected override void OnStop()
        {
            EventLog.WriteEntry("Se Finalizó el servicio windows ");
            logger.Info(DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff") + "Se Finalizó el servicio windows ");
            oConnection.Close();
            SqlDependency.Stop(connectionString);
            sendLogMail();
        }
        protected override void OnPause()
        {
            EventLog.WriteEntry("Se Pausó el servicio windows ");
            logger.Info(DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss.fff") + "Se Pausó el servicio windows ");
            sendLogMail();
        }
    }
}
