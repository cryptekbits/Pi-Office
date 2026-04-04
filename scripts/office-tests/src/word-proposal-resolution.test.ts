import assert from "node:assert/strict";
import test from "node:test";

import {
  resolveAcceptedEditParagraphIndex,
  resolveProposalParagraphIndex,
} from "../../../apps/taskpane/src/lib/office/word-proposal-resolution.js";

test("resolveProposalParagraphIndex distributes repeated search text across distinct paragraphs", () => {
  const paragraphs = [
    { index: 0, text: "Alpha repeated clause", paragraphId: "p-1" },
    { index: 1, text: "Beta repeated clause", paragraphId: "p-2" },
  ];
  const allocations = new Map<string, Set<number>>();

  const first = resolveProposalParagraphIndex(
    paragraphs,
    { kind: "replace", oldText: "repeated clause" },
    allocations,
  );
  const second = resolveProposalParagraphIndex(
    paragraphs,
    { kind: "replace", oldText: "repeated clause" },
    allocations,
  );

  assert.equal(first.index, 0);
  assert.equal(first.via, "searchText");
  assert.equal(second.index, 1);
  assert.equal(second.via, "searchText");
});

test("resolveAcceptedEditParagraphIndex prefers paragraphId over anchor and search fallback", () => {
  const paragraphs = [
    { index: 0, text: "Clause text in first paragraph.", paragraphId: "p-1" },
    { index: 1, text: "Clause text in second paragraph.", paragraphId: "p-2" },
  ];

  const resolution = resolveAcceptedEditParagraphIndex(paragraphs, {
    kind: "replace",
    paragraphId: "p-2",
    anchor: "paragraph:1",
    searchText: "Clause text",
  });

  assert.equal(resolution.index, 1);
  assert.equal(resolution.via, "paragraphId");
});

test("resolveAcceptedEditParagraphIndex prefers anchor over text-search fallback", () => {
  const paragraphs = [
    { index: 0, text: "Same clause in first paragraph.", paragraphId: "p-1" },
    { index: 1, text: "Same clause in second paragraph.", paragraphId: "p-2" },
  ];

  const resolution = resolveAcceptedEditParagraphIndex(paragraphs, {
    kind: "replace",
    anchor: "paragraph:2",
    searchText: "Same clause",
  });

  assert.equal(resolution.index, 1);
  assert.equal(resolution.via, "anchor");
});

test("resolveAcceptedEditParagraphIndex falls back to verified search text when locators do not match", () => {
  const paragraphs = [
    { index: 0, text: "Intro paragraph only.", paragraphId: "p-1" },
    { index: 1, text: "Target clause appears here.", paragraphId: "p-2" },
  ];

  const resolution = resolveAcceptedEditParagraphIndex(paragraphs, {
    kind: "replace",
    paragraphId: "missing-paragraph-id",
    anchor: "paragraph:1",
    searchText: "Target clause",
  });

  assert.equal(resolution.index, 1);
  assert.equal(resolution.via, "searchText");
});
