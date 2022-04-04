// See https://aka.ms/new-console-template for more information

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

var templatePath = Path.Join(AppDomain.CurrentDomain.BaseDirectory, "template.dotx");



var templateBytes = File.ReadAllBytes(templatePath);


using (MemoryStream templateStream = new MemoryStream())
{
    
    templateStream.Write(templateBytes, 0, (int)templateBytes.Length);

    
    using (WordprocessingDocument word = WordprocessingDocument.Open(templateStream, true))
    {

        word.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var body = word.MainDocumentPart.Document.Body;

        // create new document
        //word.AddMainDocumentPart().Document = new Document();


        // check the document settings part
        //DocumentSettingsPart docSettings = word.MainDocumentPart.DocumentSettingsPart;
        //if(docSettings == null)
        //{
        //    docSettings = AddDocumentSettings(word);
        //}



        //// check styles part
        //StyleDefinitionsPart stylePart = word.MainDocumentPart.StyleDefinitionsPart;
        //if(stylePart == null)
        //{
        //    stylePart = AddStylesPart(word);
        //}

        //// add the custom styles
        //AddHeaderStyle(word);                
        
        //AddTOC(word);
        //AddPageBreak(word);


        AddParagraph(body, "Section 1", "Heading1");
        AddParagraph(body, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse vel ultricies augue, eget vehicula massa. In at nunc lacinia, facilisis purus vitae, porttitor erat. Duis eros dui, faucibus vel imperdiet eu, fringilla ac leo. Morbi ac ex sit amet magna eleifend iaculis. Donec vel luctus felis. Etiam eros nunc, tempus pellentesque placerat vel, pulvinar quis sapien. Ut efficitur a nisl sed placerat. Mauris dignissim, dui a porta facilisis, justo sem pretium orci, id vulputate mi turpis id est. Nulla convallis nisi ac viverra tincidunt. Nulla facilisi.");
        AddParagraph(body, "Duis quis iaculis nulla. Mauris non libero porttitor, tincidunt felis sed, fringilla erat. Suspendisse facilisis sagittis erat a maximus. Pellentesque malesuada eleifend pellentesque. Phasellus malesuada, sapien rhoncus euismod convallis, metus nibh viverra elit, in cursus felis mi ut ligula. Aliquam erat volutpat. Phasellus ornare eros eget sem imperdiet imperdiet. Ut mollis eros a est interdum, vel convallis velit posuere. In ullamcorper rhoncus ante nec pellentesque. Nam eget nisi tortor. Morbi mattis, nunc maximus feugiat ornare, felis erat dictum sapien, nec ornare tellus tellus sit amet est. Proin metus metus, posuere sed sapien ut, mattis venenatis risus. Morbi aliquet pellentesque elit non semper. Donec malesuada mi non lacus eleifend dignissim. In scelerisque fringilla justo non accumsan.");


        AddParagraph(body, "Section 2", "Heading1");
        AddParagraph(body, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse vel ultricies augue, eget vehicula massa. In at nunc lacinia, facilisis purus vitae, porttitor erat. Duis eros dui, faucibus vel imperdiet eu, fringilla ac leo. Morbi ac ex sit amet magna eleifend iaculis. Donec vel luctus felis. Etiam eros nunc, tempus pellentesque placerat vel, pulvinar quis sapien. Ut efficitur a nisl sed placerat. Mauris dignissim, dui a porta facilisis, justo sem pretium orci, id vulputate mi turpis id est. Nulla convallis nisi ac viverra tincidunt. Nulla facilisi.");
        AddParagraph(body, "Duis quis iaculis nulla. Mauris non libero porttitor, tincidunt felis sed, fringilla erat. Suspendisse facilisis sagittis erat a maximus. Pellentesque malesuada eleifend pellentesque. Phasellus malesuada, sapien rhoncus euismod convallis, metus nibh viverra elit, in cursus felis mi ut ligula. Aliquam erat volutpat. Phasellus ornare eros eget sem imperdiet imperdiet. Ut mollis eros a est interdum, vel convallis velit posuere. In ullamcorper rhoncus ante nec pellentesque. Nam eget nisi tortor. Morbi mattis, nunc maximus feugiat ornare, felis erat dictum sapien, nec ornare tellus tellus sit amet est. Proin metus metus, posuere sed sapien ut, mattis venenatis risus. Morbi aliquet pellentesque elit non semper. Donec malesuada mi non lacus eleifend dignissim. In scelerisque fringilla justo non accumsan.");
        AddParagraph(body, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse vel ultricies augue, eget vehicula massa. In at nunc lacinia, facilisis purus vitae, porttitor erat. Duis eros dui, faucibus vel imperdiet eu, fringilla ac leo. Morbi ac ex sit amet magna eleifend iaculis. Donec vel luctus felis. Etiam eros nunc, tempus pellentesque placerat vel, pulvinar quis sapien. Ut efficitur a nisl sed placerat. Mauris dignissim, dui a porta facilisis, justo sem pretium orci, id vulputate mi turpis id est. Nulla convallis nisi ac viverra tincidunt. Nulla facilisi.");
        AddParagraph(body, "Duis quis iaculis nulla. Mauris non libero porttitor, tincidunt felis sed, fringilla erat. Suspendisse facilisis sagittis erat a maximus. Pellentesque malesuada eleifend pellentesque. Phasellus malesuada, sapien rhoncus euismod convallis, metus nibh viverra elit, in cursus felis mi ut ligula. Aliquam erat volutpat. Phasellus ornare eros eget sem imperdiet imperdiet. Ut mollis eros a est interdum, vel convallis velit posuere. In ullamcorper rhoncus ante nec pellentesque. Nam eget nisi tortor. Morbi mattis, nunc maximus feugiat ornare, felis erat dictum sapien, nec ornare tellus tellus sit amet est. Proin metus metus, posuere sed sapien ut, mattis venenatis risus. Morbi aliquet pellentesque elit non semper. Donec malesuada mi non lacus eleifend dignissim. In scelerisque fringilla justo non accumsan.");

        AddParagraph(body, "Section 3", "Heading1");
        AddParagraph(body, "Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Ut auctor tortor quis sem mattis, in facilisis nunc gravida. Proin eu egestas dui, vitae tristique nisl. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin facilisis ex in malesuada accumsan. Quisque blandit elit lacus, a tempus neque feugiat vel. Ut fringilla dolor eu ante ornare, eu gravida orci sodales. Fusce vel vestibulum quam. Donec nunc purus, interdum at sodales gravida, faucibus sed justo. Fusce congue, sem a convallis tristique, magna turpis faucibus lacus, ac aliquet orci lacus vitae nisl. Nunc posuere, lorem vel molestie ultricies, nulla quam lobortis ante, et finibus nibh enim ut nibh. Mauris rutrum placerat odio, eget ornare lacus tempus et. Nulla non nunc purus. Nunc ac libero non arcu vehicula sollicitudin vitae ut libero.");
        AddParagraph(body, "Vestibulum ut posuere erat, eu blandit ipsum. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer faucibus sagittis congue. Vivamus ultrices est sit amet ligula gravida, sed sollicitudin elit hendrerit. Duis id ex sed urna rhoncus porttitor. Morbi porttitor euismod turpis vel dictum. Nam tristique porta pulvinar. Etiam dictum pharetra tempus. Donec semper sem id ligula euismod, a laoreet mauris lobortis. Maecenas at est fringilla, cursus ante at, convallis nibh. Pellentesque id lacinia sapien. Nullam in elementum ex. Cras elit dolor, blandit ut eros non, placerat molestie libero.");

        AddParagraph(body, "Section 4", "Heading1");
        AddParagraph(body, "Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Ut auctor tortor quis sem mattis, in facilisis nunc gravida. Proin eu egestas dui, vitae tristique nisl. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin facilisis ex in malesuada accumsan. Quisque blandit elit lacus, a tempus neque feugiat vel. Ut fringilla dolor eu ante ornare, eu gravida orci sodales. Fusce vel vestibulum quam. Donec nunc purus, interdum at sodales gravida, faucibus sed justo. Fusce congue, sem a convallis tristique, magna turpis faucibus lacus, ac aliquet orci lacus vitae nisl. Nunc posuere, lorem vel molestie ultricies, nulla quam lobortis ante, et finibus nibh enim ut nibh. Mauris rutrum placerat odio, eget ornare lacus tempus et. Nulla non nunc purus. Nunc ac libero non arcu vehicula sollicitudin vitae ut libero.");
        AddParagraph(body, "Vestibulum ut posuere erat, eu blandit ipsum. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer faucibus sagittis congue. Vivamus ultrices est sit amet ligula gravida, sed sollicitudin elit hendrerit. Duis id ex sed urna rhoncus porttitor. Morbi porttitor euismod turpis vel dictum. Nam tristique porta pulvinar. Etiam dictum pharetra tempus. Donec semper sem id ligula euismod, a laoreet mauris lobortis. Maecenas at est fringilla, cursus ante at, convallis nibh. Pellentesque id lacinia sapien. Nullam in elementum ex. Cras elit dolor, blandit ut eros non, placerat molestie libero.");
        AddParagraph(body, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse vel ultricies augue, eget vehicula massa. In at nunc lacinia, facilisis purus vitae, porttitor erat. Duis eros dui, faucibus vel imperdiet eu, fringilla ac leo. Morbi ac ex sit amet magna eleifend iaculis. Donec vel luctus felis. Etiam eros nunc, tempus pellentesque placerat vel, pulvinar quis sapien. Ut efficitur a nisl sed placerat. Mauris dignissim, dui a porta facilisis, justo sem pretium orci, id vulputate mi turpis id est. Nulla convallis nisi ac viverra tincidunt. Nulla facilisi.");
        AddParagraph(body, "Duis quis iaculis nulla. Mauris non libero porttitor, tincidunt felis sed, fringilla erat. Suspendisse facilisis sagittis erat a maximus. Pellentesque malesuada eleifend pellentesque. Phasellus malesuada, sapien rhoncus euismod convallis, metus nibh viverra elit, in cursus felis mi ut ligula. Aliquam erat volutpat. Phasellus ornare eros eget sem imperdiet imperdiet. Ut mollis eros a est interdum, vel convallis velit posuere. In ullamcorper rhoncus ante nec pellentesque. Nam eget nisi tortor. Morbi mattis, nunc maximus feugiat ornare, felis erat dictum sapien, nec ornare tellus tellus sit amet est. Proin metus metus, posuere sed sapien ut, mattis venenatis risus. Morbi aliquet pellentesque elit non semper. Donec malesuada mi non lacus eleifend dignissim. In scelerisque fringilla justo non accumsan.");


        AddParagraph(body, "Section 5", "Heading1");
        AddParagraph(body, "Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Ut auctor tortor quis sem mattis, in facilisis nunc gravida. Proin eu egestas dui, vitae tristique nisl. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Proin facilisis ex in malesuada accumsan. Quisque blandit elit lacus, a tempus neque feugiat vel. Ut fringilla dolor eu ante ornare, eu gravida orci sodales. Fusce vel vestibulum quam. Donec nunc purus, interdum at sodales gravida, faucibus sed justo. Fusce congue, sem a convallis tristique, magna turpis faucibus lacus, ac aliquet orci lacus vitae nisl. Nunc posuere, lorem vel molestie ultricies, nulla quam lobortis ante, et finibus nibh enim ut nibh. Mauris rutrum placerat odio, eget ornare lacus tempus et. Nulla non nunc purus. Nunc ac libero non arcu vehicula sollicitudin vitae ut libero.");
        AddParagraph(body, "Vestibulum ut posuere erat, eu blandit ipsum. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer faucibus sagittis congue. Vivamus ultrices est sit amet ligula gravida, sed sollicitudin elit hendrerit. Duis id ex sed urna rhoncus porttitor. Morbi porttitor euismod turpis vel dictum. Nam tristique porta pulvinar. Etiam dictum pharetra tempus. Donec semper sem id ligula euismod, a laoreet mauris lobortis. Maecenas at est fringilla, cursus ante at, convallis nibh. Pellentesque id lacinia sapien. Nullam in elementum ex. Cras elit dolor, blandit ut eros non, placerat molestie libero.");
        AddParagraph(body, "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse vel ultricies augue, eget vehicula massa. In at nunc lacinia, facilisis purus vitae, porttitor erat. Duis eros dui, faucibus vel imperdiet eu, fringilla ac leo. Morbi ac ex sit amet magna eleifend iaculis. Donec vel luctus felis. Etiam eros nunc, tempus pellentesque placerat vel, pulvinar quis sapien. Ut efficitur a nisl sed placerat. Mauris dignissim, dui a porta facilisis, justo sem pretium orci, id vulputate mi turpis id est. Nulla convallis nisi ac viverra tincidunt. Nulla facilisi.");
        AddParagraph(body, "Duis quis iaculis nulla. Mauris non libero porttitor, tincidunt felis sed, fringilla erat. Suspendisse facilisis sagittis erat a maximus. Pellentesque malesuada eleifend pellentesque. Phasellus malesuada, sapien rhoncus euismod convallis, metus nibh viverra elit, in cursus felis mi ut ligula. Aliquam erat volutpat. Phasellus ornare eros eget sem imperdiet imperdiet. Ut mollis eros a est interdum, vel convallis velit posuere. In ullamcorper rhoncus ante nec pellentesque. Nam eget nisi tortor. Morbi mattis, nunc maximus feugiat ornare, felis erat dictum sapien, nec ornare tellus tellus sit amet est. Proin metus metus, posuere sed sapien ut, mattis venenatis risus. Morbi aliquet pellentesque elit non semper. Donec malesuada mi non lacus eleifend dignissim. In scelerisque fringilla justo non accumsan.");        

        word.MainDocumentPart.Document.Save();
    }

    using (var file = new FileStream("test.docx", FileMode.Create))
    {
        templateStream.WriteTo(file);
    }
}


/// <summary>
/// Adds the StylePart to the document
/// </summary>
StyleDefinitionsPart AddStylesPart(WordprocessingDocument document)
{
    StyleDefinitionsPart part;
    part = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
    Styles root = new Styles();
    root.Save(part);
    return part;
}


void AddPageBreak(WordprocessingDocument doc)
{
    var paragraphWithPageBreak = new Paragraph(new Run(new Break { Type = BreakValues.Page }));
    doc.MainDocumentPart.Document.Body.Append(paragraphWithPageBreak);
}

/// <summary>
/// Add "Heading" style
/// </summary>
void AddHeaderStyle(WordprocessingDocument doc)
{

    Styles styles = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

    Style headerStyle = new Style()
    {
        Type = StyleValues.Paragraph,
        StyleId = "Heading1",
        CustomStyle = false
    };

    // Create a style name
    StyleName styleName1 = new StyleName() { Val = "Heading 1" };
    BasedOn basedOn = new BasedOn() { Val = "Normal" };
    NextParagraphStyle nextParagraphStyle = new NextParagraphStyle() { Val = "Normal" };
    headerStyle.Append(styleName1);
    headerStyle.Append(basedOn);
    headerStyle.Append(nextParagraphStyle);


    // Create style properties
    StyleRunProperties styleRunProperties1 = new StyleRunProperties();
    Bold bold = new Bold();
    Color color = new Color() { ThemeColor = ThemeColorValues.Accent2 };
    RunFonts font = new RunFonts() { Ascii = "Lucida Console" };
    Italic italic = new Italic();
    FontSize fontSize = new FontSize() { Val = "24" };
    styleRunProperties1.Append(bold);
    styleRunProperties1.Append(color);
    styleRunProperties1.Append(italic);
    styleRunProperties1.Append(fontSize);


    // add the style properties to the style
    headerStyle.Append(styleRunProperties1);

    // add the style to the styles part
    styles.Append(headerStyle);

}


void AddParagraph(Body body, string text, string? styleId = null)
{
    body.AppendChild(new Paragraph());
    var paragraph = body.AppendChild(new Paragraph());

    if(!string.IsNullOrWhiteSpace(styleId))
    {
        ParagraphProperties pPr = paragraph.ParagraphProperties;
        if(pPr == null) {  pPr = new ParagraphProperties(); }

        pPr.ParagraphStyleId = new ParagraphStyleId() { Val = styleId };
        paragraph.ParagraphProperties = pPr;
    }
    
    var run = paragraph.AppendChild(new Run());
    run.AppendChild(new Text(text));
}




DocumentSettingsPart AddDocumentSettings(WordprocessingDocument doc)
{
    var docSettings = doc.MainDocumentPart.DocumentSettingsPart;
    if(docSettings == null)
    {
        docSettings = doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();        
    }

    // set the updatefieldsonopen setting -- should allow auto-updating of TOC
    docSettings.Settings = new Settings();
    docSettings.Settings.Append(new UpdateFieldsOnOpen() { Val = true });    
    docSettings.Settings.Save();

    return docSettings;
}


void AddTOC(WordprocessingDocument doc)
{
    var tocString = $@"<w:sdt>
     <w:sdtPr>
        <w:id w:val=""-493258456"" />
        <w:docPartObj>
           <w:docPartGallery w:val=""Table of Contents"" />
           <w:docPartUnique />
        </w:docPartObj>
     </w:sdtPr>
     <w:sdtEndPr>
        <w:rPr>
           <w:rFonts w:asciiTheme=""minorHAnsi"" w:eastAsiaTheme=""minorHAnsi"" w:hAnsiTheme=""minorHAnsi"" w:cstheme=""minorBidi"" />
           <w:b />
           <w:bCs />
           <w:noProof />
           <w:color w:val=""auto"" />
           <w:sz w:val=""22"" />
           <w:szCs w:val=""22"" />
        </w:rPr>
     </w:sdtEndPr>
     <w:sdtContent>
        <w:p w:rsidR=""00095C65"" w:rsidRDefault=""00095C65"">
           <w:pPr>
              <w:pStyle w:val=""CustomHeading1"" />
              <w:jc w:val=""center"" /> 
           </w:pPr>
           <w:r>
                <w:rPr>
                  <w:b /> 
                  <w:color w:val=""2E74B5"" w:themeColor=""accent1"" w:themeShade=""BF"" /> 
                  <w:sz w:val=""24"" /> 
                  <w:szCs w:val=""24"" /> 
              </w:rPr>
              <w:t>Custom Table of Contents</w:t>
           </w:r>
        </w:p>
        <w:p w:rsidR=""00095C65"" w:rsidRDefault=""00095C65"">
           <w:r>
              <w:rPr>
                 <w:b />
                 <w:bCs />
                 <w:noProof />
              </w:rPr>
              <w:fldChar w:fldCharType=""begin"" w:dirty='true' />
           </w:r>
           <w:r>
              <w:rPr>
                 <w:b />
                 <w:bCs />
                 <w:noProof />
              </w:rPr>
              <w:instrText xml:space=""preserve""> TOC \o ""1-3"" \h \z \u </w:instrText>
           </w:r>
           <w:r>
              <w:rPr>
                 <w:b />
                 <w:bCs />
                 <w:noProof />
              </w:rPr>
              <w:fldChar w:fldCharType=""separate"" />
           </w:r>
           <w:r>
              <w:rPr>
                 <w:noProof />
              </w:rPr>
              <w:t>No table of contents entries found.</w:t>
           </w:r>
           <w:r>
              <w:rPr>
                 <w:b />
                 <w:bCs />
                 <w:noProof />
              </w:rPr>
              <w:fldChar w:fldCharType=""end"" />
           </w:r>
        </w:p>
     </w:sdtContent>
  </w:sdt>";


    var sdtBlock = new SdtBlock();
    sdtBlock.InnerXml = tocString;
    doc.MainDocumentPart.Document.Body.AppendChild(sdtBlock);
    doc.MainDocumentPart.Document.Save();
    
}