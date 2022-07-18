  SELECT
  Documents.DocumentID,
  Documents.Filename,
  Documents.Deleted,
  Projects.Path
FROM Documents
JOIN DocumentsInProjects
  ON Documents.DocumentID = DocumentsInProjects.DocumentID
JOIN Projects
  ON Projects.ProjectID = DocumentsInProjects.ProjectID

  where Documents.Filename IN (SELECT Filename FROM Documents GROUP BY Filename HAVING count(*) >=2) AND Documents.Deleted = 0 AND Projects.Path like '\other\Data\%' Order by Filename