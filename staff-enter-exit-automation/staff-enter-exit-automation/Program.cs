using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace staffenterexitautomation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            decimal overtimeBonus = 0;
            int salary = 50;

            // Get employee's entering and exiting time
            Console.WriteLine("Enter the employee's entering time(AM): ");
            string enteringTime = Console.ReadLine();
            Console.WriteLine("Enter the employee's exiting time(PM): ");
            string exitingTime = Console.ReadLine();

            // Convert the entering and exiting time to DateTime
            DateTime entering = DateTime.Parse(enteringTime);
            DateTime exiting = DateTime.Parse(exitingTime);

            // Check if the employee worked overtime and calculate the overtime bonus
            if (exiting.Hour > 17)
            {
                TimeSpan overtime = exiting - new DateTime(exiting.Year, exiting.Month, exiting.Day, 17, 0, 0);
                overtimeBonus = overtime.Hours * salary / 2;
                Console.WriteLine("Overtime bonus: " + overtimeBonus);
            }

            // Calculate the total hours worked and the total salary
            TimeSpan totalHours = exiting - entering;
            decimal totalSalary = totalHours.Hours * salary + overtimeBonus;

            // Print the results
            Console.WriteLine($"Exit time: {exiting.Hour}");
            Console.WriteLine("Total hours worked: " + totalHours.Hours);
            Console.WriteLine("Total salary: " + totalSalary);
            Console.ReadLine();
        }
    }
}
