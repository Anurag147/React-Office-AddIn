/* global Word console */

export async function insertText(text: string) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export const highlightPara = async (sentenceToHighlight: string) => {
  // Run this code in the context of a Word add-in
  Word.run(function (context) {
    // Specify the sentence to highlight
    //var sentenceToHighlight = "This is an example sentence.";

    // Search for all occurrences of the sentence
    var searchResults = context.document.body.search(sentenceToHighlight, { matchCase: false, matchWholeWord: false });

    // Load the search results
    context.load(searchResults, "text");

    // Synchronize the document state by running the queued commands
    return context
      .sync()
      .then(function () {
        // Loop through the search results and highlight each occurrence
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.highlightColor = "#FFFF00"; // Yellow highlight
        }
      })
      .then(context.sync);
  }).catch(function (error) {
    console.log("Error: " + error.message);
  });
};
