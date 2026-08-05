/// <summary>
/// A template written to the driver's temp directory before the generator runs.
/// </summary>
/// <param name="FileName">
/// Path relative to that directory. May be nested, in which case the folder is created.
/// </param>
/// <param name="Bytes">The file's content — markdown as UTF-8, or a built docx.</param>
sealed record TemplateFile(string FileName, byte[] Bytes);
