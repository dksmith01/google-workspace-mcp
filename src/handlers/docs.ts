import type { drive_v3, docs_v1 } from "googleapis";
import {
  log,
  structuredResponse,
  errorResponse,
  withTimeout,
  validateArgs,
} from "../utils/index.js";
import { GOOGLE_MIME_TYPES, getMimeTypeSuggestion } from "../utils/mimeTypes.js";
import type { ToolResponse } from "../utils/index.js";
import {
  CreateGoogleDocSchema,
  UpdateGoogleDocSchema,
  GetGoogleDocContentSchema,
  AppendToDocSchema,
  InsertTextInDocSchema,
  DeleteTextInDocSchema,
  ReplaceTextInDocSchema,
  FormatGoogleDocRangeSchema,
} from "../schemas/index.js";
import { resolveOptionalFolderPath, checkFileExists } from "./helpers.js";
import { toDocsColorStyle } from "../utils/colors.js";

/**
 * Get the end index of a Google Doc's content.
 * Used for calculating document length and insert positions.
 */
function getDocumentEndIndex(document: docs_v1.Schema$Document): number {
  const content = document.body?.content;
  if (!content || content.length === 0) return 1;
  return content[content.length - 1]?.endIndex || 1;
}

// -----------------------------------------------------------------------------
// FRONTMATTER BOX STYLING
// -----------------------------------------------------------------------------
// Drive's markdown import renders a leading fenced code block (typically YAML
// frontmatter) as plain default-font paragraphs with blank paragraphs between
// lines. These helpers restore a Proof-style box: monospace light text on a
// solid dark shaded block with padding.

const FRONTMATTER_TEXT_COLOR = { red: 0.92, green: 0.92, blue: 0.92 };
const FRONTMATTER_BOX_COLOR = { red: 0.05, green: 0.05, blue: 0.05 };
const FRONTMATTER_PADDING_PT = 8;

function findLeadingFenceLines(content: string): string[] | null {
  const match = content.match(/^```[a-zA-Z]*\r?\n([\s\S]*?)\r?\n```(?:\r?\n|$)/);
  if (!match) return null;
  const lines = match[1].split(/\r?\n/).filter((line) => line.trim() !== "");
  return lines.length > 0 ? lines : null;
}

interface ParagraphInfo {
  startIndex: number;
  endIndex: number;
  text: string;
}

function getParagraphs(document: docs_v1.Schema$Document): ParagraphInfo[] {
  const paragraphs: ParagraphInfo[] = [];
  for (const element of document.body?.content ?? []) {
    if (!element.paragraph) continue;
    const text = (element.paragraph.elements ?? [])
      .map((textElement) => textElement.textRun?.content ?? "")
      .join("");
    paragraphs.push({
      startIndex: element.startIndex ?? 0,
      endIndex: element.endIndex ?? 0,
      text,
    });
  }
  return paragraphs;
}

interface FenceLocation {
  blankRanges: Array<{ startIndex: number; endIndex: number }>;
  endIndex: number;
}

/** Match the fence lines against leading doc paragraphs, collecting import-artifact blanks. */
function locateImportedFenceLines(
  paragraphs: ParagraphInfo[],
  lines: string[],
): FenceLocation | null {
  const blankRanges: FenceLocation["blankRanges"] = [];
  let endIndex = 0;
  let p = 0;
  for (const line of lines) {
    while (p < paragraphs.length && paragraphs[p].text === "\n") {
      blankRanges.push({
        startIndex: paragraphs[p].startIndex,
        endIndex: paragraphs[p].endIndex,
      });
      p++;
    }
    if (p >= paragraphs.length || paragraphs[p].text.trimEnd() !== line.trimEnd()) return null;
    endIndex = paragraphs[p].endIndex;
    p++;
  }
  return { blankRanges, endIndex };
}

function buildFrontmatterBoxRequests(location: FenceLocation): docs_v1.Schema$Request[] {
  const requests: docs_v1.Schema$Request[] = [];
  const descending = [...location.blankRanges].sort((a, b) => b.startIndex - a.startIndex);
  for (const range of descending) {
    requests.push({ deleteContentRange: { range } });
  }

  const deletedChars = location.blankRanges.reduce(
    (total, range) => total + (range.endIndex - range.startIndex),
    0,
  );
  const range = { startIndex: 1, endIndex: location.endIndex - deletedChars };

  requests.push({
    updateTextStyle: {
      range,
      textStyle: {
        weightedFontFamily: { fontFamily: "Courier New" },
        fontSize: { magnitude: 10, unit: "PT" },
        foregroundColor: toDocsColorStyle(FRONTMATTER_TEXT_COLOR),
      },
      fields: "weightedFontFamily,fontSize,foregroundColor",
    },
  });

  const border = {
    color: {},
    dashStyle: "SOLID",
    padding: { magnitude: FRONTMATTER_PADDING_PT, unit: "PT" },
    width: { magnitude: 0, unit: "PT" },
  };
  requests.push({
    updateParagraphStyle: {
      range,
      paragraphStyle: {
        shading: { backgroundColor: toDocsColorStyle(FRONTMATTER_BOX_COLOR) },
        spaceAbove: { magnitude: 0, unit: "PT" },
        spaceBelow: { magnitude: 0, unit: "PT" },
        borderLeft: border,
        borderRight: border,
        borderTop: border,
        borderBottom: border,
      },
      fields:
        "shading.backgroundColor,spaceAbove,spaceBelow," +
        "borderLeft,borderRight,borderTop,borderBottom",
    },
  });
  return requests;
}

/**
 * Style a markdown-imported leading fence as a boxed block.
 * Never throws: the doc write already succeeded, so styling failures are
 * logged and reported via the returned flag instead of failing the tool call.
 */
async function styleImportedFrontmatter(
  docs: docs_v1.Docs,
  documentId: string,
  content: string,
): Promise<boolean> {
  const lines = findLeadingFenceLines(content);
  if (!lines) return false;

  try {
    const document = await docs.documents.get({ documentId });
    const location = locateImportedFenceLines(getParagraphs(document.data), lines);
    if (!location) {
      log("Frontmatter styling skipped: imported paragraphs did not match fence lines", {
        documentId,
      });
      return false;
    }
    await docs.documents.batchUpdate({
      documentId,
      requestBody: { requests: buildFrontmatterBoxRequests(location) },
    });
    return true;
  } catch (styleError) {
    log("Frontmatter styling failed", { documentId, error: String(styleError) });
    return false;
  }
}

export async function handleCreateGoogleDoc(
  drive: drive_v3.Drive,
  docs: docs_v1.Docs,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(CreateGoogleDocSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  const parentFolderId = await resolveOptionalFolderPath(
    drive,
    data.parentFolderId,
    data.parentPath,
  );

  // Check if document already exists
  const existingFileId = await checkFileExists(drive, data.name, parentFolderId);
  if (existingFileId) {
    return errorResponse(
      `A document named "${data.name}" already exists in this location. ` +
        `To update it, use updateGoogleDoc with documentId: ${existingFileId}`,
      { code: "ALREADY_EXISTS", context: { existingFileId } },
    );
  }

  log("Creating Google Doc", { parentFolderId, contentFormat: data.contentFormat });

  // Markdown content is converted natively by Drive at creation time
  const useMarkdownImport = data.contentFormat === "markdown" && data.content.length > 0;

  let docResponse;
  try {
    docResponse = await drive.files.create({
      requestBody: {
        name: data.name,
        mimeType: "application/vnd.google-apps.document",
        parents: [parentFolderId],
      },
      ...(useMarkdownImport && {
        media: { mimeType: "text/markdown", body: data.content },
      }),
      fields: "id, name, webViewLink",
      supportsAllDrives: true,
    });
  } catch (createError: unknown) {
    const err = createError as {
      message?: string;
      code?: number;
      errors?: unknown;
      status?: number;
    };
    log("Drive files.create error details:", {
      message: err.message,
      code: err.code,
      errors: err.errors,
      status: err.status,
    });
    throw createError;
  }
  const doc = docResponse.data;

  let frontmatterStyled: boolean | undefined;
  if (useMarkdownImport && data.styleFrontmatter) {
    frontmatterStyled = await styleImportedFrontmatter(docs, doc.id!, data.content);
  }

  if (!useMarkdownImport && data.content.length > 0) {
    await docs.documents.batchUpdate({
      documentId: doc.id!,
      requestBody: {
        requests: [
          {
            insertText: { location: { index: 1 }, text: data.content },
          },
          // Ensure the text is formatted as normal text, not as a header
          {
            updateParagraphStyle: {
              range: {
                startIndex: 1,
                endIndex: data.content.length + 1,
              },
              paragraphStyle: {
                namedStyleType: "NORMAL_TEXT",
              },
              fields: "namedStyleType",
            },
          },
        ],
      },
    });
  }

  return structuredResponse(
    `Created Google Doc: ${doc.name}\nID: ${doc.id}\nLink: ${doc.webViewLink}`,
    {
      id: doc.id!,
      name: doc.name!,
      webViewLink: doc.webViewLink!,
      ...(frontmatterStyled !== undefined && { frontmatterStyled }),
    },
  );
}

export async function handleUpdateGoogleDoc(
  drive: drive_v3.Drive,
  docs: docs_v1.Docs,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(UpdateGoogleDocSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  if (data.contentFormat === "markdown") {
    // Drive converts markdown natively, replacing the doc's full content
    const updateResponse = await drive.files.update({
      fileId: data.documentId,
      media: { mimeType: "text/markdown", body: data.content },
      fields: "name",
      supportsAllDrives: true,
    });

    let frontmatterStyled: boolean | undefined;
    if (data.styleFrontmatter) {
      frontmatterStyled = await styleImportedFrontmatter(docs, data.documentId, data.content);
    }

    return structuredResponse(`Updated Google Doc: ${updateResponse.data.name}`, {
      title: updateResponse.data.name!,
      updated: true,
      ...(frontmatterStyled !== undefined && { frontmatterStyled }),
    });
  }

  const document = await docs.documents.get({ documentId: data.documentId });

  // Delete all content
  const endIndex = getDocumentEndIndex(document.data);

  // Google Docs API doesn't allow deleting the final newline character
  // We need to leave at least one character in the document
  const deleteEndIndex = Math.max(1, endIndex - 1);

  if (deleteEndIndex > 1) {
    await docs.documents.batchUpdate({
      documentId: data.documentId,
      requestBody: {
        requests: [
          {
            deleteContentRange: {
              range: { startIndex: 1, endIndex: deleteEndIndex },
            },
          },
        ],
      },
    });
  }

  // Insert new content
  await docs.documents.batchUpdate({
    documentId: data.documentId,
    requestBody: {
      requests: [
        {
          insertText: { location: { index: 1 }, text: data.content },
        },
        // Ensure the text is formatted as normal text, not as a header
        {
          updateParagraphStyle: {
            range: {
              startIndex: 1,
              endIndex: data.content.length + 1,
            },
            paragraphStyle: {
              namedStyleType: "NORMAL_TEXT",
            },
            fields: "namedStyleType",
          },
        },
      ],
    },
  });

  return structuredResponse(`Updated Google Doc: ${document.data.title}`, {
    title: document.data.title!,
    updated: true,
  });
}

export async function handleGetGoogleDocContent(
  drive: drive_v3.Drive,
  docs: docs_v1.Docs,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(GetGoogleDocContentSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  // Check file type before calling Docs API to provide helpful error messages
  const metadata = await drive.files.get({
    fileId: data.documentId,
    fields: "mimeType,name",
    supportsAllDrives: true,
  });

  const mimeType = metadata.data.mimeType;
  if (mimeType !== GOOGLE_MIME_TYPES.DOCUMENT) {
    const fileName = metadata.data.name || data.documentId;
    const suggestion = getMimeTypeSuggestion(mimeType);
    return errorResponse(`"${fileName}" is not a Google Doc (type: ${mimeType}). ${suggestion}`);
  }

  if (data.format === "markdown") {
    const exportResponse = await withTimeout(
      drive.files.export(
        { fileId: data.documentId, mimeType: "text/markdown" },
        { responseType: "text" },
      ),
      30000,
      "Export document as markdown",
    );
    const markdown = String(exportResponse.data ?? "");

    return structuredResponse(markdown, {
      documentId: data.documentId,
      title: metadata.data.name,
      markdown,
      totalLength: markdown.length,
    });
  }

  const document = await withTimeout(
    docs.documents.get({ documentId: data.documentId }),
    30000,
    "Get document content",
  );

  const contentSegments: Array<{
    startIndex: number;
    endIndex: number;
    text: string;
  }> = [];
  let content = "";
  let currentIndex = 1;

  // Extract text content with indices
  if (document.data.body?.content) {
    for (const element of document.data.body.content) {
      if (element.paragraph?.elements) {
        for (const textElement of element.paragraph.elements) {
          if (textElement.textRun?.content) {
            const text = textElement.textRun.content;
            const startIdx = currentIndex;
            content += text;
            currentIndex += text.length;
            contentSegments.push({
              startIndex: startIdx,
              endIndex: currentIndex,
              text: text,
            });
          }
        }
      }
    }
  }

  // Format the response to show text with indices
  let formattedContent = "Document content with indices:\n\n";
  let lineStart = 1;
  const lines = content.split("\n");

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineEnd = lineStart + line.length;
    if (line.trim()) {
      formattedContent += `[${lineStart}-${lineEnd}] ${line}\n`;
    }
    lineStart = lineEnd + 1; // +1 for the newline character
  }

  const textResponse = formattedContent + `\nTotal length: ${content.length} characters`;

  return structuredResponse(textResponse, {
    documentId: data.documentId,
    title: document.data.title,
    content: contentSegments,
    totalLength: content.length,
  });
}

export async function handleAppendToDoc(docs: docs_v1.Docs, args: unknown): Promise<ToolResponse> {
  const validation = validateArgs(AppendToDocSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  // Get document to find end index
  const document = await docs.documents.get({ documentId: data.documentId });
  const endIndex = getDocumentEndIndex(document.data);

  // Insert at end index - 1 (before the final newline)
  const insertIndex = Math.max(1, endIndex - 1);

  // Prepare the text to insert
  const textToInsert = data.insertNewline ? `\n${data.text}` : data.text;

  await docs.documents.batchUpdate({
    documentId: data.documentId,
    requestBody: {
      requests: [
        {
          insertText: {
            location: { index: insertIndex },
            text: textToInsert,
          },
        },
      ],
    },
  });

  log("Text appended to document", {
    documentId: data.documentId,
    textLength: data.text.length,
  });

  return structuredResponse(
    `Appended ${data.text.length} characters to document "${document.data.title}"`,
    {
      title: document.data.title!,
      charactersAdded: data.text.length,
    },
  );
}

export async function handleInsertTextInDoc(
  docs: docs_v1.Docs,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(InsertTextInDocSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  // Get document to validate index and get title
  const document = await docs.documents.get({ documentId: data.documentId });
  const docLength = getDocumentEndIndex(document.data);

  if (data.index >= docLength) {
    return errorResponse(
      `Index ${data.index} is beyond the document length (${docLength - 1} characters). ` +
        `Use appendToDoc to add text at the end, or specify an index between 1 and ${docLength - 1}.`,
    );
  }

  await docs.documents.batchUpdate({
    documentId: data.documentId,
    requestBody: {
      requests: [
        {
          insertText: {
            location: { index: data.index },
            text: data.text,
          },
        },
      ],
    },
  });

  log("Text inserted into document", {
    documentId: data.documentId,
    index: data.index,
    textLength: data.text.length,
  });

  return structuredResponse(
    `Inserted ${data.text.length} characters at index ${data.index} in "${document.data.title}"`,
    {
      title: document.data.title!,
      index: data.index,
      charactersInserted: data.text.length,
    },
  );
}

export async function handleDeleteTextInDoc(
  docs: docs_v1.Docs,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(DeleteTextInDocSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  // Get document to validate indices and get title
  const document = await docs.documents.get({ documentId: data.documentId });
  const docLength = getDocumentEndIndex(document.data);

  if (data.endIndex > docLength) {
    return errorResponse(
      `End index ${data.endIndex} is beyond the document length (${docLength - 1} characters). ` +
        `Valid range is 1 to ${docLength - 1}.`,
    );
  }

  const charsToDelete = data.endIndex - data.startIndex;

  await docs.documents.batchUpdate({
    documentId: data.documentId,
    requestBody: {
      requests: [
        {
          deleteContentRange: {
            range: {
              startIndex: data.startIndex,
              endIndex: data.endIndex,
            },
          },
        },
      ],
    },
  });

  log("Text deleted from document", {
    documentId: data.documentId,
    startIndex: data.startIndex,
    endIndex: data.endIndex,
  });

  return structuredResponse(
    `Deleted ${charsToDelete} characters (indices ${data.startIndex}-${data.endIndex}) from "${document.data.title}"`,
    {
      title: document.data.title!,
      startIndex: data.startIndex,
      endIndex: data.endIndex,
      charactersDeleted: charsToDelete,
    },
  );
}

export async function handleReplaceTextInDoc(
  docs: docs_v1.Docs,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(ReplaceTextInDocSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  // Get document title
  const document = await docs.documents.get({ documentId: data.documentId });

  const response = await docs.documents.batchUpdate({
    documentId: data.documentId,
    requestBody: {
      requests: [
        {
          replaceAllText: {
            containsText: {
              text: data.searchText,
              matchCase: data.matchCase,
            },
            replaceText: data.replaceText,
          },
        },
      ],
    },
  });

  // Get the number of replacements made
  const occurrencesChanged = response.data.replies?.[0]?.replaceAllText?.occurrencesChanged || 0;

  log("Text replaced in document", {
    documentId: data.documentId,
    occurrences: occurrencesChanged,
  });

  if (occurrencesChanged === 0) {
    return structuredResponse(
      `No occurrences of "${data.searchText}" found in "${document.data.title}"`,
      {
        title: document.data.title!,
        occurrencesChanged: 0,
      },
    );
  }

  return structuredResponse(
    `Replaced ${occurrencesChanged} occurrence(s) of "${data.searchText}" with "${data.replaceText}" in "${document.data.title}"`,
    {
      title: document.data.title!,
      occurrencesChanged,
    },
  );
}

// -----------------------------------------------------------------------------
// DOC FORMATTING HELPERS
// -----------------------------------------------------------------------------

interface DocTextFormatOptions {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  foregroundColor?: { red?: number; green?: number; blue?: number };
  backgroundColor?: { red?: number; green?: number; blue?: number };
}

function buildDocTextStyle(data: DocTextFormatOptions): {
  style: Record<string, unknown>;
  fields: string[];
} {
  const style: Record<string, unknown> = {};
  const fields: string[] = [];

  if (data.bold !== undefined) {
    style.bold = data.bold;
    fields.push("bold");
  }
  if (data.italic !== undefined) {
    style.italic = data.italic;
    fields.push("italic");
  }
  if (data.underline !== undefined) {
    style.underline = data.underline;
    fields.push("underline");
  }
  if (data.strikethrough !== undefined) {
    style.strikethrough = data.strikethrough;
    fields.push("strikethrough");
  }
  if (data.fontSize !== undefined) {
    style.fontSize = { magnitude: data.fontSize, unit: "PT" };
    fields.push("fontSize");
  }
  if (data.fontFamily !== undefined) {
    style.weightedFontFamily = { fontFamily: data.fontFamily };
    fields.push("weightedFontFamily");
  }
  if (data.foregroundColor) {
    style.foregroundColor = toDocsColorStyle(data.foregroundColor);
    fields.push("foregroundColor");
  }
  if (data.backgroundColor) {
    style.backgroundColor = toDocsColorStyle(data.backgroundColor);
    fields.push("backgroundColor");
  }

  return { style, fields };
}

interface DocParagraphFormatOptions {
  namedStyleType?: string;
  paragraphBackgroundColor?: { red?: number; green?: number; blue?: number };
  paragraphPadding?: number;
  alignment?: string;
  lineSpacing?: number;
  spaceAbove?: number;
  spaceBelow?: number;
}

function buildDocParagraphStyle(data: DocParagraphFormatOptions): {
  style: Record<string, unknown>;
  fields: string[];
} {
  const style: Record<string, unknown> = {};
  const fields: string[] = [];

  if (data.namedStyleType !== undefined) {
    style.namedStyleType = data.namedStyleType;
    fields.push("namedStyleType");
  }
  if (data.alignment !== undefined) {
    style.alignment = data.alignment;
    fields.push("alignment");
  }
  if (data.lineSpacing !== undefined) {
    style.lineSpacing = data.lineSpacing;
    fields.push("lineSpacing");
  }
  if (data.spaceAbove !== undefined) {
    style.spaceAbove = { magnitude: data.spaceAbove, unit: "PT" };
    fields.push("spaceAbove");
  }
  if (data.spaceBelow !== undefined) {
    style.spaceBelow = { magnitude: data.spaceBelow, unit: "PT" };
    fields.push("spaceBelow");
  }
  if (data.paragraphBackgroundColor) {
    // Shading fills the full paragraph width, unlike text backgroundColor (glyph highlight)
    style.shading = { backgroundColor: toDocsColorStyle(data.paragraphBackgroundColor) };
    fields.push("shading.backgroundColor");
  }
  if (data.paragraphPadding !== undefined) {
    // Invisible zero-width borders carry the padding; shading extends into the padded area
    const border = {
      color: {},
      dashStyle: "SOLID",
      padding: { magnitude: data.paragraphPadding, unit: "PT" },
      width: { magnitude: 0, unit: "PT" },
    };
    style.borderLeft = border;
    style.borderRight = border;
    style.borderTop = border;
    style.borderBottom = border;
    fields.push("borderLeft", "borderRight", "borderTop", "borderBottom");
  }

  return { style, fields };
}

// -----------------------------------------------------------------------------
// DOC FORMATTING HANDLER
// -----------------------------------------------------------------------------

export async function handleFormatGoogleDocRange(
  docs: docs_v1.Docs,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(FormatGoogleDocRangeSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  const document = await docs.documents.get({ documentId: data.documentId });
  const docEndIndex = getDocumentEndIndex(document.data);
  const startIndex = data.startIndex ?? 1;
  const endIndex = data.endIndex ?? docEndIndex;

  const requests: docs_v1.Schema$Request[] = [];
  const formatsApplied: string[] = [];

  const textResult = buildDocTextStyle(data);
  if (textResult.fields.length > 0) {
    requests.push({
      updateTextStyle: {
        range: { startIndex, endIndex },
        textStyle: textResult.style,
        fields: textResult.fields.join(","),
      },
    });
    formatsApplied.push(...textResult.fields);
  }

  const paragraphResult = buildDocParagraphStyle(data);
  if (paragraphResult.fields.length > 0) {
    requests.push({
      updateParagraphStyle: {
        range: { startIndex, endIndex },
        paragraphStyle: paragraphResult.style,
        fields: paragraphResult.fields.join(","),
      },
    });
    formatsApplied.push(...paragraphResult.fields);
  }

  if (requests.length === 0) {
    return errorResponse(
      "No formatting options specified. Provide at least one of: " +
        "bold, italic, underline, strikethrough, fontSize, fontFamily, foregroundColor, " +
        "backgroundColor, paragraphBackgroundColor, paragraphPadding, namedStyleType, " +
        "alignment, lineSpacing, spaceAbove, spaceBelow.",
    );
  }

  await docs.documents.batchUpdate({
    documentId: data.documentId,
    requestBody: { requests },
  });

  log("Applied formatting to document range", {
    documentId: data.documentId,
    startIndex,
    endIndex,
    formatsApplied,
  });
  return structuredResponse(
    `Applied formatting to range ${startIndex}-${endIndex}: ${formatsApplied.join(", ")}`,
    {
      startIndex,
      endIndex,
      formatsApplied,
    },
  );
}
