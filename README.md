# MD2ML
Convert Markdown text in WordProcessingML using OpenXML specifications

# Usage Instructions:

    using (Md2Ml.Md2MlEngine engine = new Md2Ml.Md2MlEngine())
    {
      engine.CreateDocument(@"D:\DOCX-Playground\Empty.docx", @"D:\DOCX-PlayGround\NewDocumentMd2Ml.docx");
      engine.WriteMdText("# Intro\nGo ahead, play around with the editor! Be sure to check out **bold** and *italic* styling, or even [links](https://google.com). You can type the Markdown syntax, use the toolbar, or use shortcuts like `cmd-b` or `ctrl-b`.\n\n## Lists\nUnordered lists can be started using the toolbar or by typing `*`, `-`, or `+`. Ordered lists can be started by typing `1. `.\n\n### Unordered\n* Lists are a piece of cake\n* They even auto continue as you type\n* A double enter will end them\n* Tabs and shift-tabs work too\n\n### Ordered\n1. Numbered lists...\n2. ...work too!\n\n## What about images?\n![Yes](https://i.imgur.com/sZlktY7.png)\n| Ashok | Arun RD | Himalaya |\n|:-|:-:|-:|\n|As | Ar | 12 |\n|As | Ar | 12 |");
    }
    
# Extended Usage Instructions:
Apart from writing Markdown Text to Worddocument this document also let you draft Word Document from the ground.

| Function | Description |
|---|---|
|`CreateParagraph`| Creates a Paragraph attach it to `Document`'s body |
|`CreateNonBodyParagraph`| Create a Paragraph without attaching it to `Document`. Is this useful for creating tables, pagebreaks etc.,|
|`WriteText`| Writes a plain text to document |
|`WriteMdText`| Writes Markdown formatted text to document |
|`WriteWebsiteHyperlink`| Add Hyperlink |
|`NumberedList`| Adds ordered list to document (Plain text) |
|`MarkdownNumberedList`| Adds ordered list to document (Markdown formatted text) |
|`BulletedList`| Adds unordered list to document (Plain text) |
|`MarkdownBulletedList`|Adds unordered list to document (Markdown formatted text) |
|`AddPageBreak`| Inserts Pagebreak |
|`StartBookmark`| Starts Bookmark regions (Useful for creating TOC and Hyperlinks) |
|`EndBookmark`| Ends Bookmark regions |
|`CreateTable`| Create a Reference to `Table` object |
|`AddTableRow`| Write table cells (Non Formatted) |
|`AddImage`| Insert image to document from specified URL or file location |
|`Cleanup`| Clear entire document |
|`SaveDocument`| Save documnet |

