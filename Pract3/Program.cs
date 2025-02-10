using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Linq;


namespace Pract3
{
    public enum lect_names
    {
        Теория_вероятностей,
        Теория_вычислимости,
        Геометрическое_моделирование,
        КСЕ,
        Программирование,
        Алгоритмы,
        ООП
    }
    public enum pract_names
    {
        Программирование,
        Алгоритмы,
        ООП
    }
    public class Classroom
    {
        public object number;
        public string type;

        public Classroom() { }
        public Classroom(object num, string type)
        {
            this.type = type;
            this.number = num;
        }

        public int Set_number
        {
            set { number = value; }
        }
        public string Type
        {
            set { type = value; }
        }

    }
    public class Subject
    {
        public object name;
        public Subject() { }
        public Subject(object name) => this.name = name;

        public string Name
        {
            set { name = value; }
        }
    }
    public class Grup
    {
        public string number;
        public Grup() { }
        public Grup(string number) => this.number = number;
        public string num
        {
            set { number = value; }
        }
    }

    public class Lesson
    {
        public Classroom classroom;
        public int para;
        public Subject subject;
        public Grup grup;
        public int day;
        public Lesson(object num, string type, string grup_num, object name, int para, int day)
        {
            classroom = new Classroom(num, type);
            subject = new Subject(name);
            grup = new Grup(grup_num);
            this.para = para;
            this.day = day;
        }

    }
    public class Week_Timetable
    {
        public int count_grup = 0;
        public int count_subjects = 0;
        public int count_lect = 0;
        public int count_pract = 0;
        public int count_per_day = 0;
        public int count_lectori = 0;
        public int count_terminal = 0;
        public lect_names[] lectia_names = (lect_names[])Enum.GetValues(typeof(lect_names));
        public pract_names[] practica_names = (pract_names[])Enum.GetValues(typeof(pract_names));
        public List<Lesson> lessons = new List<Lesson>();
        public List<string> grup = new List<string>();
        public Dictionary<System.Collections.Generic.KeyValuePair<int, int>, bool> lectoria_para =
            new Dictionary<System.Collections.Generic.KeyValuePair<int, int>, bool>();
        public Dictionary<System.Collections.Generic.KeyValuePair<int, int>, bool> terminal_para =
            new Dictionary<System.Collections.Generic.KeyValuePair<int, int>, bool>();
        public int ost_lect = 0;
        public int ost_pract = 0;
        public int been_in_voen = 1;
        public Week_Timetable(int count_grup, int count_subject, int count_lect,
            int count_pract, int count_per_day, int count_lectori, int count_terminal)
        {
            this.count_lect = count_lect;
            this.count_per_day = count_per_day;
            this.count_pract = count_pract;
            this.count_lectori = count_lectori;
            this.count_grup = count_grup;
            this.count_subjects = count_subject;
            this.count_terminal = count_terminal;
            for (int i = 1; i <= count_grup; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    grup.Add(i.ToString());
                }
            }

            for (int i = 0; i < count_lectori; i++)
            {
                int num_lectori = i + 1;
                for (int j = 0; j < 6; j++)
                {
                    System.Collections.Generic.KeyValuePair<int, int> lect_para =
                        new System.Collections.Generic.KeyValuePair<int, int>(num_lectori, j + 1);

                    lectoria_para[lect_para] = true;
                }
            }

            for (int i = 0; i < count_terminal; i++)
            {
                int num_terminal = i + 1;
                for (int j = 0; j < 6; j++)
                {
                    System.Collections.Generic.KeyValuePair<int, int> term_para =
                        new System.Collections.Generic.KeyValuePair<int, int>(num_terminal, j + 1);

                    terminal_para[term_para] = true;
                }
            }
            ost_lect = count_lect;
            ost_pract = count_pract;

        }
        public void Refresh_auditoris_for_new_day()
        {
            for (int i = 0; i < lectoria_para.Count(); i++)
            {
                var k_ey = lectoria_para.ElementAt(i).Key;
                lectoria_para[k_ey] = true;
            }
            for (int i = 0; i < terminal_para.Count(); i++)
            {
                var k_ey = terminal_para.ElementAt(i).Key;
                terminal_para[k_ey] = true;
            }
        }
        public void Print()
        {
            for (int i = 0; i < lessons.Count(); i++)
            {
                Console.WriteLine("ДЕНЬ ");
                Console.WriteLine(lessons[i].day);
                Console.WriteLine("-----------------");
                Console.WriteLine("Номер пары");
                Console.WriteLine(lessons[i].para.ToString());
                Console.WriteLine("Номер аудитории ");
                Console.WriteLine(lessons[i].classroom.number.ToString());
                Console.WriteLine("Номер группы ");
                Console.WriteLine(lessons[i].grup.number);
                Console.WriteLine("Предмет ");
                Console.WriteLine(lessons[i].subject.name);
                Console.WriteLine("Тип занятия ");
                Console.WriteLine(lessons[i].classroom.type);
            }
            Console.ReadLine();
        }
        public void Create_timetable()
        {
            for (int i = 0; i < 5; i++)
            {
                Day_timetable(i + 1);
                Refresh_auditoris_for_new_day();
            }
        }
        public void Day_timetable(int day)
        {
            int g = 0;
            int flag = 0;
            int count_lect_for_grup = count_per_day;
            Dictionary<int, HashSet<int>> group_pars = new Dictionary<int, HashSet<int>>();
            if (this.ost_lect > 0)
            {
                for (int i = 0; i < grup.Count(); i++)
                {
                    if (group_pars.ContainsKey(i + 1) == true && group_pars[i + 1].Count == count_per_day) break;
                    if (i + 1 != been_in_voen)
                    {
                        if (ost_pract != 0)
                        {
                            double need_for_pract = ost_pract / 5;
                            if (need_for_pract < 1) need_for_pract = 1;
                            else need_for_pract = (int)Math.Ceiling(need_for_pract);
                            count_lect_for_grup -= ((int)need_for_pract);
                        }
                        HashSet<int> group_been_in_pars = new HashSet<int>();
                        for (int j = 0; j < count_lect_for_grup; j++)
                        {

                            while (true)
                            {
                                Random random = new Random();
                                int k = random.Next(0, lectoria_para.Count() - 1);
                                var k_ey = lectoria_para.ElementAt(k).Key;
                                if (!group_been_in_pars.Contains(k_ey.Value) && lectoria_para[k_ey] == true)
                                {
                                    Lesson lesson = new Lesson(k_ey.Key, "Лекция", (i + 1).ToString(), lectia_names[g], k_ey.Value, day);
                                    if (g == count_subjects - 1)
                                    {
                                        g = 0;
                                    }
                                    else g++;
                                    lectoria_para[k_ey] = false;
                                    ost_lect--;
                                    lessons.Add(lesson);
                                    group_been_in_pars.Add(k_ey.Value);
                                    break;
                                }
                                else
                                {
                                    if (lectoria_para[k_ey] == false) flag++;
                                }
                                if (flag == lectoria_para.Count() || ost_lect == 0) break;

                            }
                            if (ost_lect == 0) break;
                        }
                        if (group_been_in_pars.Count() != 0) group_pars[i + 1] = group_been_in_pars;
                    }
                    if (ost_lect == 0) break;
                }
                Lesson voenka = new Lesson("Гагарина", "Военная кафедра", (been_in_voen).ToString(), " ", 1, day);
                lessons.Add(voenka);
            }
            if (this.ost_pract > 0)
            {
                double need_for_pract = ost_pract / 5;
                if (need_for_pract < 1) need_for_pract = 1;
                else need_for_pract = (int)Math.Ceiling(need_for_pract);
                for (int i = 0; i < grup.Count(); i++)
                {
                    if (group_pars.ContainsKey(i + 1) == true && group_pars[i + 1].Count == count_per_day) break;
                    bool create = false;
                    if (i + 1 != been_in_voen)
                    {
                        for (int j = 0; j < need_for_pract; j++)
                        {

                            while (true)
                            {
                                Random random = new Random();
                                int k = random.Next(0, terminal_para.Count() - 1);
                                var k_ey = terminal_para.ElementAt(k).Key;
                                if (terminal_para[k_ey] == true)
                                {
                                    if (group_pars.ContainsKey(i + 1) == false)
                                    {
                                        HashSet<int> group_been_in_pars = new HashSet<int>();
                                        group_been_in_pars.Add(k_ey.Value);
                                        group_pars[i + 1] = group_been_in_pars;
                                        create = true;
                                    }

                                    if (create == true || (!group_pars[i + 1].Contains(k_ey.Value)))
                                    {
                                        Lesson lesson = new Lesson(k_ey.Key, "Практика", (i + 1).ToString(), practica_names[g], k_ey.Value, day);
                                        if ((g > count_subjects && g == count_subjects - 1) || (g < count_subjects && g == practica_names.Count() - 1))
                                        {
                                            g = 0;
                                        }
                                        else g++;
                                        terminal_para[k_ey] = false;
                                        ost_pract--;
                                        lessons.Add(lesson);
                                        group_pars[i + 1].Add(k_ey.Value);
                                        break;
                                    }
                                }
                                else
                                {
                                    flag++;
                                }
                                if (flag == terminal_para.Count() || ost_pract == 0) break;

                            }
                            if (ost_pract == 0) break;

                        }
                    }
                    if (ost_pract == 0) break;

                }
            }
            been_in_voen += 1;

        }

        public void Convert_to_Exel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("WeekTimetable");
                var filePath = "C:/Users/armok/OneDrive/Рабочий стол";
                string[] daysOfWeek = { "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота" };
                for (int i = 0; i < daysOfWeek.Length; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = daysOfWeek[i];

                    for (int j = 0; j < 6; j++) // 6 пар максимум
                    {
                        worksheet.Cells[i + 2, j + 2].Value = "";
                    }
                }
                foreach (var lesson in lessons)
                {
                    int dayIndex = lesson.day - 1; // Индекс дня недели
                    int pairIndex = 0; // Индекс пары (можно настроить по необходимости)

                    // Поиск первой свободной ячейки для данного дня
                    for (int j = 0; j < 6; j++)
                    {
                        if (string.IsNullOrEmpty(worksheet.Cells[dayIndex + 2, j + 2].Text))
                        {
                            pairIndex = j;
                            break;
                        }
                        worksheet.Cells[dayIndex + 2, pairIndex + 2].Value = $"{lesson.classroom.number}\n{lesson.classroom.type}\n{lesson.grup.number}\n{lesson.subject.name}";
                        worksheet.Cells[dayIndex + 2, pairIndex + 2].Style.WrapText = true;
                    }
                }
                var fileInfo = new FileInfo(filePath);
                excelPackage.SaveAs(fileInfo);
            }

        }

    }
    class Program
    {
        static void Main()
        {
            Week_Timetable a = new Week_Timetable(4, 5, 20, 5, 4, 3, 2);
            a.Create_timetable();
            a.Convert_to_Exel();

        }
    }
}


