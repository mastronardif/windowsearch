using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("ZZZZZZZZZZZZZ Hello World AAAAAAAAAAAAAAAA");
            search11();
            //bobo();
        }

        static void search11()
        {
            OleDbConnection conn = new OleDbConnection("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';");

            conn.Open();

            string strQry = m_Qry;
            OleDbCommand cmd = new OleDbCommand(strQry, conn);
            //doAsaReaderObj(cmd);
            doAsaRecordSet(cmd);
            conn.Close();
        }

        static string m_Qry =
            //"SELECT top 10 " +
            "SELECT  " +
         "System.ItemPathDisplay " +
         //",System.DocTitle " + // find an equivalent
         ",System.ItemName " +
         ",System.Size " +
         ",System.Search.Rank " +
         ",System.ItemUrl " +
         ",System.ItemPathDisplayNarrow " +
         ",System.FileName " +
        "FROM SYSTEMINDEX" +
       " WHERE (SCOPE='file:C:\\FxM\\Dev\\joeschedule\\cgi\\ngfop' " +
             " or SCOPE='file:C:\\Users\\gp' " +
       //" or SCOPE='file:C:\\Users\\frmastronardi' " +
       ") " +
       " and contains(*, 'go') order by System.Search.Rank desc" +
            "";

        static void doAsaRecordSet(OleDbCommand cmd)
        {
            // RecordSet
            var objConnection = new ADODB.Connection();
            var rs = new ADODB.Recordset();

            objConnection.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';");
            //rs.Open(m_Qry, objConnection);
            // rs.Open(sql, cnnDBreport, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic);

            rs.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly;
            //rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient; //adUseServer;// adUseClient;
            rs.Open(m_Qry, objConnection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic);
            

            Console.WriteLine(rs);

            rs.MoveFirst();

            int iCnt = 0;
            string col = "System.FileName";
            //while (rs.EOF == false)
            for (int iRow = 1; rs.EOF == false; iRow++)
            {
                Console.WriteLine("{0}) {1}", ++iCnt, rs.Fields[col].Value);
                rs.MoveNext();

                if (10 == iRow)
                {
                    Console.WriteLine("{0}) AbsolutePage= {1} rs.AbsolutePosition= {2}", iRow, rs.AbsolutePosition, rs.RecordCount);
                    //rs.CursorLocation
                    //rs.Fields
                    //rs.MaxRecords
                    //rs.PageCount
                    //rs.PageSize
                    //rs.Properties
                    



                    Console.WriteLine("Hit Enter Key.");
                    Console.ReadLine();
                    iRow = 0;
                }

            }

            rs.Close();

        }

        static void doAsaReaderObj(OleDbCommand cmd)
        {
            OleDbDataReader rdr = null;
            rdr = cmd.ExecuteReader();

            //RS.AbsolutePage
            int iCnt = 0;
            while (rdr.Read())
            {
                string col = "System.Search.Rank";
                Console.Write("{0}) {1}: {2} ", ++iCnt, col, rdr[rdr.GetOrdinal(col)]);                

                col = "System.ItemPathDisplay";
                Console.Write("{0}: {1} ", col, rdr[rdr.GetOrdinal(col)]);
                col = "System.Size";
                Console.WriteLine("{0}: {1} ", col, rdr[rdr.GetOrdinal(col)]);
            }

            rdr.Close();
        }

        static void bobo()
        {
            using (OleDbConnection conn = new OleDbConnection(
    "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand("office", conn);

                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    //gridResults.Rows.Clear();

                    //gridResults.Columns.Clear();

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        //gridResults.Columns.Add(reader.GetName(i), reader.GetName(i));
                        Console.WriteLine(reader.GetName(i), reader.GetOrdinal("System.Size"));
                    }

                    while (reader.Read())
                    {
                        List<object> row = new List<object>();

                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row.Add(reader[i]);
                        }

                        //gridResults.Rows.Add(row.ToArray());
                        Console.WriteLine(row.ToArray());
                    }
                }
            }

        }

    }
}
