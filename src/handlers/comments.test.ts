import { describe, it, expect, vi, beforeEach } from "vitest";
import type { drive_v3 } from "googleapis";
import { handleListComments, handleReplyToComment, handleResolveComment } from "./comments.js";

vi.mock("../utils/index.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../utils/index.js")>();
  return {
    ...actual,
    log: vi.fn(),
  };
});

function createMockDrive(): drive_v3.Drive {
  return {
    comments: { list: vi.fn() },
    replies: { create: vi.fn() },
  } as unknown as drive_v3.Drive;
}

const openComment = {
  id: "comment1",
  content: "Please fix this paragraph",
  quotedFileContent: { value: "the paragraph text" },
  author: { displayName: "David Smith" },
  createdTime: "2026-07-23T10:00:00Z",
  modifiedTime: "2026-07-23T10:00:00Z",
  resolved: false,
  replies: [
    {
      id: "reply1",
      content: "Working on it",
      author: { displayName: "Claude" },
      createdTime: "2026-07-23T11:00:00Z",
    },
  ],
};

const resolvedComment = {
  id: "comment2",
  content: "Old feedback",
  author: { displayName: "David Smith" },
  createdTime: "2026-07-20T10:00:00Z",
  resolved: true,
  replies: [],
};

describe("handleListComments", () => {
  let mockDrive: drive_v3.Drive;

  beforeEach(() => {
    mockDrive = createMockDrive();
  });

  it("lists open comments with quoted text and replies", async () => {
    vi.mocked(mockDrive.comments.list).mockResolvedValue({
      data: { comments: [openComment, resolvedComment] },
    } as never);

    const result = await handleListComments(mockDrive, { fileId: "file123" });
    expect(result.isError).toBe(false);
    expect(result.content[0].text).toContain("Please fix this paragraph");
    expect(result.content[0].text).toContain("the paragraph text");
    expect(result.content[0].text).toContain("Working on it");
    expect(result.content[0].text).not.toContain("Old feedback");
  });

  it("includes resolved comments when includeResolved is true", async () => {
    vi.mocked(mockDrive.comments.list).mockResolvedValue({
      data: { comments: [openComment, resolvedComment] },
    } as never);

    const result = await handleListComments(mockDrive, {
      fileId: "file123",
      includeResolved: true,
    });
    expect(result.isError).toBe(false);
    expect(result.content[0].text).toContain("Old feedback");
    expect(result.content[0].text).toContain("[resolved]");
  });

  it("follows pagination across pages", async () => {
    vi.mocked(mockDrive.comments.list)
      .mockResolvedValueOnce({
        data: { comments: [openComment], nextPageToken: "page2" },
      } as never)
      .mockResolvedValueOnce({
        data: { comments: [{ ...openComment, id: "comment3", content: "Second page" }] },
      } as never);

    const result = await handleListComments(mockDrive, { fileId: "file123" });
    expect(result.isError).toBe(false);
    expect(mockDrive.comments.list).toHaveBeenCalledTimes(2);
    expect(result.content[0].text).toContain("Second page");
  });

  it("handles file with no comments", async () => {
    vi.mocked(mockDrive.comments.list).mockResolvedValue({
      data: { comments: [] },
    } as never);

    const result = await handleListComments(mockDrive, { fileId: "file123" });
    expect(result.isError).toBe(false);
    expect(result.content[0].text).toContain("No open comments");
  });

  it("returns error for empty fileId", async () => {
    const result = await handleListComments(mockDrive, { fileId: "" });
    expect(result.isError).toBe(true);
  });
});

describe("handleReplyToComment", () => {
  let mockDrive: drive_v3.Drive;

  beforeEach(() => {
    mockDrive = createMockDrive();
  });

  it("creates a reply on the comment", async () => {
    vi.mocked(mockDrive.replies.create).mockResolvedValue({
      data: { id: "reply42" },
    } as never);

    const result = await handleReplyToComment(mockDrive, {
      fileId: "file123",
      commentId: "comment1",
      content: "Done, see latest version",
    });
    expect(result.isError).toBe(false);
    expect(result.content[0].text).toContain("Replied to comment");
    expect(mockDrive.replies.create).toHaveBeenCalledWith(
      expect.objectContaining({
        fileId: "file123",
        commentId: "comment1",
        requestBody: { content: "Done, see latest version" },
      }),
    );
  });

  it("returns error for empty content", async () => {
    const result = await handleReplyToComment(mockDrive, {
      fileId: "file123",
      commentId: "comment1",
      content: "",
    });
    expect(result.isError).toBe(true);
  });
});

describe("handleResolveComment", () => {
  let mockDrive: drive_v3.Drive;

  beforeEach(() => {
    mockDrive = createMockDrive();
  });

  it("resolves via a reply with resolve action", async () => {
    vi.mocked(mockDrive.replies.create).mockResolvedValue({
      data: { id: "reply43", action: "resolve" },
    } as never);

    const result = await handleResolveComment(mockDrive, {
      fileId: "file123",
      commentId: "comment1",
    });
    expect(result.isError).toBe(false);
    expect(result.content[0].text).toContain("Resolved comment");
    expect(mockDrive.replies.create).toHaveBeenCalledWith(
      expect.objectContaining({
        requestBody: { action: "resolve" },
      }),
    );
  });

  it("includes closing reply content when provided", async () => {
    vi.mocked(mockDrive.replies.create).mockResolvedValue({
      data: { id: "reply44", action: "resolve" },
    } as never);

    const result = await handleResolveComment(mockDrive, {
      fileId: "file123",
      commentId: "comment1",
      content: "Addressed in latest revision",
    });
    expect(result.isError).toBe(false);
    expect(mockDrive.replies.create).toHaveBeenCalledWith(
      expect.objectContaining({
        requestBody: { action: "resolve", content: "Addressed in latest revision" },
      }),
    );
  });

  it("returns error for empty commentId", async () => {
    const result = await handleResolveComment(mockDrive, {
      fileId: "file123",
      commentId: "",
    });
    expect(result.isError).toBe(true);
  });
});
