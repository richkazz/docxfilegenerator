using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxFileGenerator.Models;

namespace DoxcFileGenerator.Utils;

public static class ResumeToWordDocument
{
    public static void GenerateWordDocument(Resume resume, string outputPath)
    {
        using WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = mainPart.Document.AppendChild(new Body());

        if (resume.Name is not null)
        {
            // Add name
            AddHeading(body, resume.Name, 1);
        }

        // Add contact information
        AddParagraph(body, $"{resume.Address}\n{resume.Email}\n{resume.PhoneNumber}");

        if (resume.Objective is not null)
        {
            // Add objective
            AddHeading(body, "OBJECTIVE", 2);
            AddParagraph(body, resume.Objective);
        }
        // Add projects
        AddHeading(body, "PROJECTS", 2);
        foreach (var project in resume.Projects)
        {
            AddHeading(body, $"Product Name: {project.Name}", 3);
            AddParagraph(body, $"Technologies: {string.Join(", ", project.Technologies)}");
            AddParagraph(body, $"Description: {project.Description}");

            if (project.KeyFeatures.Count != 0)
            {
                AddParagraph(body, "Key Features:");
                AddBulletList(body, project.KeyFeatures);
            }
        }

        // Add work experience
        AddHeading(body, "WORK EXPERIENCE", 2);
        foreach (var experience in resume.WorkExperience)
        {
            AddHeading(body, experience.Company, 3);
            AddParagraph(body, $"Position: {experience.Position}");
            AddBulletList(body, experience.Responsibilities);
        }

        // Add profile
        AddHeading(body, "PROFILE", 2);
        if (resume.Profile is not null)
        {
            AddBulletList(body, resume.Profile.Split('\n').Select(s => s.Trim()).ToList());
        }

        // Add education
        AddHeading(body, "EDUCATION", 2);
        foreach (var edu in resume.Education)
        {
            AddParagraph(body, $"{edu.Institution}\n{edu.Degree}");
        }

        // Add skills
        AddHeading(body, "SKILLS", 2);
        AddParagraph(body, string.Join(", ", resume.Skills));

        // Add certifications
        AddHeading(body, "CERTIFICATION", 2);
        foreach (var cert in resume.Certifications)
        {
            var year = cert.Date is null ? "N/A" : cert.Date.GetValueOrDefault().Year.ToString();
            AddParagraph(body, $"{cert.Name} - {cert.Organization} {year}");
            AddParagraph(body, $"Verify online: {cert.VerificationLink}");
        }

        // Add contact information
        AddHeading(body, "CONTACT INFORMATION:", 2);
        AddParagraph(body, $"LinkedIn: {resume.ContactInformation.LinkedIn}");

        // Add referee
        AddHeading(body, "REFEREE", 2);
        AddParagraph(body, resume.Referee);
    }

    private static void AddHeading(Body body, string text, int level)
    {
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        RunProperties runProperties = run.AppendChild(new RunProperties());
        runProperties.AppendChild(new Bold());
        runProperties.AppendChild(new FontSize() { Val = (32 - (level - 1) * 4).ToString() });
        run.AppendChild(new Text(text));
    }

    private static void AddParagraph(Body body, string text)
    {
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(text));
    }

    private static void AddBulletList(Body body, List<string> items)
    {
        foreach (var item in items)
        {
            Paragraph para = body.AppendChild(new Paragraph());
            ParagraphProperties paraProp = para.AppendChild(new ParagraphProperties());
            NumberingProperties numberingProps = paraProp.AppendChild(new NumberingProperties());
            numberingProps.AppendChild(new NumberingLevelReference() { Val = 0 });
            numberingProps.AppendChild(new NumberingId() { Val = 1 });
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(item));
        }
    }
}