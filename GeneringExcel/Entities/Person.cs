using Microsoft.EntityFrameworkCore;

namespace Entities;

public class Person
{
    public int Id { get; set; }
    public string Name { get; set; } = null!;
    [Precision(12, 2)]
    public decimal Salary { get; set; }
    public DateTime DateOfBirth { get; set; }
}