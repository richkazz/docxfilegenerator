using System;
using System.Collections.Generic;

namespace DocxFileGenerator.Models;
public class Resume
{
    public string? Name { get; set; }
    public string? Address { get; set; }
    public string? PhoneNumber { get; set; }
    public string? Email { get; set; }
    public string? Objective { get; set; }
    public List<string> Skills { get; set; } = [];
    public List<WorkExperience> WorkExperience { get; set; } = [];
    public List<Education> Education { get; set; } = [];
    public List<Project> Projects { get; set; } = [];
    public string? Profile { get; set; }
    public List<Certification> Certifications { get; set; } = new List<Certification>();
    public ContactInformation ContactInformation { get; set; }
    public string Referee { get; set; }
}

public class WorkExperience
{
    public string Company { get; set; }
    public string Position { get; set; }
    public DateTime? StartDate { get; set; }
    public DateTime? EndDate { get; set; }
    public List<string> Responsibilities { get; set; } = new List<string>();
}

public class Education
{
    public string Institution { get; set; }
    public string Degree { get; set; }
    public DateTime? GraduationDate { get; set; }
}

public class Project
{
    public string Name { get; set; }
    public List<string> Technologies { get; set; } = new List<string>();
    public string Description { get; set; }
    public List<string> KeyFeatures { get; set; } = new List<string>();
}

public class Certification
{
    public string Name { get; set; }
    public string Organization { get; set; }
    public DateTime? Date { get; set; }
    public string VerificationLink { get; set; }
}

public class ContactInformation
{
    public string LinkedIn { get; set; }
}