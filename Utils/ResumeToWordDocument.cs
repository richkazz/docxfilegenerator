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

public static class ResumeToWordDocumentTemplateTwo
{
    public static void GenerateWordDocument(Resume resume, string outputPath)
    {
        using WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = mainPart.Document.AppendChild(new Body());

        // Add name and contact info
        AddCenteredParagraph(body, resume.Name ?? "", true, "36");
        AddCenteredParagraph(body, resume.Address ?? "", false, "24");
        AddCenteredParagraph(body, resume.Email ?? "", false, "24");
        AddCenteredParagraph(body, resume.PhoneNumber ?? "", false, "24");

        // Add objective
        AddHeading(body, "OBJECTIVE", "28");
        AddParagraph(body, resume.Objective ?? "");

        // Add projects
        AddHeading(body, "PROJECTS", "28");
        foreach (var project in resume.Projects)
        {
            AddSubHeading(body, $"Product Name: {project.Name}");
            AddBulletPoint(body, $"Technologies: {string.Join(", ", project.Technologies)}");
            AddBulletPoint(body, $"Description: {project.Description}");

            if (project.KeyFeatures.Count != 0)
            {
                AddBulletPoint(body, "Key Features:");
                foreach (var feature in project.KeyFeatures)
                {
                    AddBulletPoint(body, feature, 1);
                }
            }
        }

        // Add work experience
        AddHeading(body, "WORK EXPERIENCE", "28");
        foreach (var experience in resume.WorkExperience)
        {
            AddSubHeading(body, experience.Company);
            AddParagraph(body, $"Position: {experience.Position}");
            foreach (var responsibility in experience.Responsibilities)
            {
                AddBulletPoint(body, responsibility);
            }
        }

        // Add profile
        if (resume.Profile is not null)
        {
            AddHeading(body, "PROFILE", "28");
            foreach (var profileItem in resume.Profile.Split('\n').Select(s => s.Trim()))
            {
                AddBulletPoint(body, profileItem);
            }
        }


        // Add education
        AddHeading(body, "EDUCATION", "28");
        foreach (var edu in resume.Education)
        {
            AddBulletPoint(body, $"{edu.Institution}");
            AddBulletPoint(body, $"{edu.Degree}", 1);
        }

        // Add skills
        AddHeading(body, "SKILLS", "28");
        AddParagraph(body, string.Join(", ", resume.Skills));

        // Add certifications
        AddHeading(body, "CERTIFICATION", "28");
        foreach (var cert in resume.Certifications)
        {
            var year = cert.Date is null ? "N/A" : cert.Date.GetValueOrDefault().Year.ToString();
            AddBulletPoint(body, $"{cert.Name} - {cert.Organization} {year}");
            if (!string.IsNullOrWhiteSpace(cert.VerificationLink))
            {
                AddHyperlink(mainPart, body, "Verify online", cert.VerificationLink);
            }
        }

        // Add contact information
        AddHeading(body, "CONTACT INFORMATION:", "28");
        AddParagraph(body, $"LinkedIn: ");
        if (!string.IsNullOrWhiteSpace(resume.ContactInformation.LinkedIn))
        {
            AddHyperlink(mainPart, body, resume.ContactInformation.LinkedIn, resume.ContactInformation.LinkedIn);
        }
        // Add referee
        AddHeading(body, "REFEREE", "28");
        AddParagraph(body, resume.Referee);
    }

    private static void AddHeading(Body body, string text, string fontSize)
    {
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        RunProperties runProperties = run.AppendChild(new RunProperties());
        runProperties.AppendChild(new Bold());
        runProperties.AppendChild(new FontSize() { Val = fontSize });
        run.AppendChild(new Text(text));
    }

    private static void AddSubHeading(Body body, string text)
    {
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        RunProperties runProperties = run.AppendChild(new RunProperties());
        runProperties.AppendChild(new Bold());
        runProperties.AppendChild(new FontSize() { Val = "24" });
        run.AppendChild(new Text(text));
    }

    private static void AddParagraph(Body body, string text)
    {
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(text));
    }

    private static void AddCenteredParagraph(Body body, string text, bool isBold, string fontSize)
    {
        Paragraph para = body.AppendChild(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center })));
        Run run = para.AppendChild(new Run());
        RunProperties runProperties = run.AppendChild(new RunProperties());
        if (isBold) runProperties.AppendChild(new Bold());
        runProperties.AppendChild(new FontSize() { Val = fontSize });
        run.AppendChild(new Text(text));
    }

    private static void AddBulletPoint(Body body, string text, int level = 0)
    {
        Paragraph para = body.AppendChild(new Paragraph());
        ParagraphProperties paraProp = para.AppendChild(new ParagraphProperties());
        NumberingProperties numberingProps = paraProp.AppendChild(new NumberingProperties());
        numberingProps.AppendChild(new NumberingLevelReference() { Val = level });
        numberingProps.AppendChild(new NumberingId() { Val = 1 });
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(text));
    }

    private static void AddHyperlink(MainDocumentPart mainPart, Body body, string text, string url)
    {
        // Add hyperlink relationship
        HyperlinkRelationship hyperlinkRelationship = mainPart.AddHyperlinkRelationship(new Uri(url), true);
        string relationshipId = hyperlinkRelationship.Id;

        // Create the hyperlink
        Hyperlink hyperlink = new() { Id = relationshipId, History = OnOffValue.FromBoolean(true) };
        Run hyperlinkRun = new(new Text(text));
        hyperlink.Append(hyperlinkRun);

        // Add the hyperlink to the document
        Paragraph hyperlinkParagraph = new(hyperlink);
        body.Append(hyperlinkParagraph);
    }
}