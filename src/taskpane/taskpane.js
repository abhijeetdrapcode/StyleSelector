Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("logStyleContentButton").onclick = logStyleContent;
    document.getElementById("displayAllStylesButton").onclick = displayAllStyles;
    document.getElementById("displayHeadingsWithNumberingButton").onclick = displayHeadingsWithNumbering;
    document.getElementById("displayAllListsButton").onclick = getListInfo;
    document.getElementById("getListInfoByStyleButton").onclick = getListInfoByStyle;
  }
});

async function logStyleContent() {
  const styleName = document.getElementById("styleInput").value.trim();

  if (!styleName) {
    console.error("Please enter a style name.");
    return;
  }

  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style,items/text");

    await context.sync();

    let matchingParagraphs = "";

    paragraphs.items.forEach((paragraph, index) => {
      if (paragraph.style === styleName) {
        console.log(`Paragraph ${index + 1}: Style - ${paragraph.style}, Text - "${paragraph.text}"`);
        matchingParagraphs += `Paragraph ${index + 1}: ${paragraph.text}\n`;
      }
    });

    if (!matchingParagraphs) {
      console.log(`No content found with style "${styleName}".`);
    }

    await context.sync();
  }).catch(errorHandler);
}

async function displayAllStyles() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style,items/text");

    await context.sync();

    paragraphs.items.forEach((paragraph, index) => {
      console.log(`Paragraph ${index + 1}: Style - ${paragraph.style}, Text - "${paragraph.text}"`);
    });

    await context.sync();
  }).catch(errorHandler);
}

async function displayHeadingsWithNumbering() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style,items/text,items/listItemInfo");

    await context.sync();

    let headingDetails = "";

    paragraphs.items.forEach((paragraph, index) => {
      const style = paragraph.style;
      const listItemInfo = paragraph.listItemInfo;
      const paragraphText = paragraph.text.trim();

      console.log(style);
      if (style && style.toLowerCase().includes("heading") && listItemInfo && listItemInfo.levelString) {
        const numbering = listItemInfo.levelString;
        headingDetails += `Heading ${index + 1}: ${numbering} - ${paragraphText}\n`;
      }
    });

    if (headingDetails) {
      console.log("Headings with Numbering:\n" + headingDetails);
    } else {
      console.log("No headings with numbering found.");
    }

    await context.sync();
  }).catch(errorHandler);
}

async function getListInfo() {
  try {
    await Word.run(async (context) => {
      console.log("Word.run started");

      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      console.log(`Total paragraphs in the document: ${paragraphs.items.length}`);

      let currentList = [];
      let currentLevel = -1;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text");
        await context.sync();

        try {
          paragraph.load("isListItem");
          await context.sync();

          if (paragraph.isListItem) {
            paragraph.listItem.load("level,listString");
            await context.sync();

            const level = paragraph.listItem.level;
            const listString = paragraph.listItem.listString || "";
            const text = paragraph.text.trim();

            if (level <= currentLevel && currentList.length > 0) {
              console.log(currentList.join("\n"));
              currentList = [];
            }

            const indent = "  ".repeat(level);
            currentList.push(`${indent}${listString} ${text}`);
            currentLevel = level;
          } else if (currentList.length > 0) {
            console.log(currentList.join("\n"));
            currentList = [];
            currentLevel = -1;
          }
        } catch (error) {
          if (currentList.length > 0) {
            console.log(currentList.join("\n"));
            currentList = [];
            currentLevel = -1;
          }
          console.error(`Error processing paragraph ${i}:`, error);
        }
      }

      if (currentList.length > 0) {
        console.log(currentList.join("\n"));
      }
    });
  } catch (error) {
    console.error("An error occurred:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
  }
}

async function getListInfoByStyle() {
  const styleName = document.getElementById("styleInput").value.trim();

  if (!styleName) {
    console.error("Please enter a style name.");
    return;
  }

  try {
    await Word.run(async (context) => {
      console.log(`Getting list info for style: ${styleName}`);

      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      console.log(`Total paragraphs in the document: ${paragraphs.items.length}`);

      let currentList = [];
      let currentLevel = -1;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,style");
        await context.sync();

        if (paragraph.style === styleName) {
          try {
            paragraph.load("isListItem");
            await context.sync();

            if (paragraph.isListItem) {
              paragraph.listItem.load("level,listString");
              await context.sync();

              const level = paragraph.listItem.level;
              const listString = paragraph.listItem.listString || "";
              const text = paragraph.text.trim();

              if (level <= currentLevel && currentList.length > 0) {
                console.log(currentList.join("\n"));
                currentList = [];
              }

              const indent = "  ".repeat(level);
              console.log(indent);
              console.log(listString);
              currentList.push(`${indent}${listString} ${text}`);
              currentLevel = level;
            } else {
              const text = paragraph.text.trim();
              currentList.push(text);
            }
          } catch (error) {
            const text = paragraph.text.trim();
            currentList.push(text);
            console.error(`Error processing paragraph ${i}:`, error);
          }
        } else if (currentList.length > 0) {
          console.log(currentList.join("\n"));
          currentList = [];
          currentLevel = -1;
        }
      }

      if (currentList.length > 0) {
        console.log(currentList.join("\n"));
      }
    });
  } catch (error) {
    console.error("An error occurred:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
  }
}

function errorHandler(error) {
  console.error("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.error("Debug info: " + JSON.stringify(error.debugInfo));
  }
}
