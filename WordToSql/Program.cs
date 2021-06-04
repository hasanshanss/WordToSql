using Spire.Doc;
using Spire.Doc.Collections;
using Spire.Doc.Documents;
using System;
using System.Dynamic;
using System.IO;
using System.Text;

namespace WordToSql
{
    class Program
    {

        static void Main(string[] args)
        {

            Document doc = new Document("1.docx");
            Section section = doc.Sections[0];



            //Table table = section.Tables[1] as Table;

            foreach (Table table in section.Tables)
            {
                StringBuilder query = new StringBuilder("INSERT INTO `belediyyeler` (`name`, `officer`, `address`, `email`, `phone`, `working_days`, `working_hours`,  `members`, `servants`, `participants`) VALUES");
                StringBuilder paragraphs = new StringBuilder();
                string[] data = new string[10];

                Console.OutputEncoding = System.Text.Encoding.UTF8;

                table.Rows.Remove(table.Rows.FirstItem);
                int rows = 0;
                foreach (TableRow row in table.Rows)
                {
                    int i = 0;
                    foreach (TableCell cell in row.Cells)
                    {
                        paragraphs.Append("'");

                        if (i == 9 && rows > 0)
                        {
                            try
                            {

                                foreach (TableRow tableRow in cell.Tables[0].Rows)
                                {

                                    foreach (TableCell tableCell in tableRow.Cells)
                                    {
                                        foreach (Paragraph tableP in tableCell.Paragraphs)
                                        {
                                            paragraphs.Append($"{tableP.Text} ");

                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {

                            }
                            paragraphs.Append("'");
                            data[i] = paragraphs.ToString();
                            paragraphs.Clear();
                            //string join1 = string.Join(",", data);
                            //Console.WriteLine(join1);
                            //query.AppendLine($" ({join1}),");
                            rows++;
                            continue;
                        }

                        foreach (Paragraph p in cell.Paragraphs)
                        {
                            if (i == 5)
                            {
                                paragraphs.Append($"{p.Text} ");
                                paragraphs.Append("'");
                                data[i] = paragraphs.ToString();
                                paragraphs.Clear();
                                paragraphs.Append("'");
                                i++;
                                continue;
                            }

                       

                            if (p.Text.Equals("yoxdur"))
                            {
                                p.Text = "0";
                            }

                            //if (i == row.Cells.Count - 1 || i == row.Cells.Count)
                            //{
                            //    if (!Int32.TryParse(p.Text.Trim(), out int x))
                            //    {
                            //        //if (cell.Paragraphs.Count > 1)
                            //        //{
                            //        //    if(cell.Paragraphs[0] == p)
                            //        //    {
                            //        //        continue;
                            //        //    }
                            //        //}
                            //        p.Text = "0";
                            //    }
                            //}

                            paragraphs.Append($"{p.Text} ");
                        }
                        paragraphs.Append("'");

                        data[i] = paragraphs.ToString();
                        i++;

                        paragraphs.Clear();
                    }
                    string join = string.Join(",", data);
                    Console.WriteLine(join);
                    query.AppendLine($" ({join}),");
                    rows++;
                }
                //query.Replace(',',';',query.Length-2,1);
                using (System.IO.StreamWriter file = new System.IO.StreamWriter("1.sql", true))
                {
                    file.WriteLine(query);
                }
            }
        }

    }
}

