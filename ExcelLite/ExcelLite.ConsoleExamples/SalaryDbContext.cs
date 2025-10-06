using Microsoft.EntityFrameworkCore;

namespace ExcelLite.ConsoleExamples
{
    public class SalaryDbContext : DbContext
    {
        public SalaryDbContext(DbContextOptions<SalaryDbContext> options)
        : base(options)
        {
        }

        public DbSet<DbSalary> Salaries => Set<DbSalary>();
    }
}
