function fillSlidesWithJson() {
  const jsonData = {
    "questions": [
      {
        "year": 1992,
        "question": "Which album was released by Nirvana in 1992?",
        "options": ["In Utero", "Bleach", "Nevermind", "Incesticide"],
        "correct_answer": "Incesticide",
        "fact": "\"Incesticide\" is a compilation of rare and demo recordings released by Nirvana in December 1992.",
        "source": "https://en.wikipedia.org/wiki/Incesticide"
      },
      {
        "year": 1993,
        "question": "Who released the album \"Music Box\" in 1993?",
        "options": ["Whitney Houston", "Mariah Carey", "Madonna", "Janet Jackson"],
        "correct_answer": "Mariah Carey",
        "fact": "\"Music Box\" was Mariah Carey's third studio album and included hits like \"Dreamlover\" and \"Hero\".",
        "source": "https://en.wikipedia.org/wiki/Music_Box_(Mariah_Carey_album)"
      }
      // Add more if needed
    ]
  };

  const PREFIXES = {
    year: "",
    question: "",
    variant1: "1: ",
    variant2: "2: ",
    variant3: "3: ",
    variant4: "4: ",
    fact: "",
    source: "src: ",
    answer: ""
  };

  const CORRECT_ANSWER_COLOR = "#77eb7e";

  const presentation = SlidesApp.getActivePresentation();
  const templateSlide = presentation.getSlides()[0];

  for (let i = jsonData.questions.length - 1; i >= 0; i--) {
    const questionData = jsonData.questions[i];
    Logger.log(`Creating slides for question: "${questionData.question}" (Slide number: ${i + 1})`);
    createAnswerSlide(templateSlide, questionData, PREFIXES, CORRECT_ANSWER_COLOR, true, i + 1); // Pass true to remove the answer figure
    createQuestionSlide(templateSlide, questionData, PREFIXES, i + 1);
  }
}

// Function to replace text in shapes
function replaceTextInShape(element, searchText, replaceText) {
  const textRange = element.asShape().getText();
  const currentText = textRange.asString();
  if (currentText.toLowerCase().includes(searchText.toLowerCase())) {
    textRange.replaceAllText(searchText, replaceText);
  }
}

// Function to replace shape with an image
function replaceShapeWithImage(slide, element, imageUrl) {
  try {
    const blob = UrlFetchApp.fetch(imageUrl).getBlob();
    const position = element.getLeft();
    const top = element.getTop();
    const width = element.getWidth();
    const height = element.getHeight();
    slide.insertImage(blob, position, top, width, height);
    element.remove();
  } catch (error) {
    Logger.log(`Error replacing shape with image: ${error.toString()}`);
  }
}

// Function to download image URL
function getPreviewImageUrl(url) {
  try {
    let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      throw new Error('Network response was not ok');
    }
    let text = response.getContentText();

    let ogImage = text.match(/<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i);
    if (ogImage && ogImage[1]) {
      Logger.log(`Found og:image: ${ogImage[1]}`);
      return ogImage[1];
    }

    let twitterImage = text.match(/<meta[^>]*name=["']twitter:image["'][^>]*content=["']([^"']+)["']/i);
    if (twitterImage && twitterImage[1]) {
      Logger.log(`Found twitter:image: ${twitterImage[1]}`);
      return twitterImage[1];
    }

    let linkImage = text.match(/<link[^>]*rel=["']image_src["'][^>]*href=["']([^"']+)["']/i);
    if (linkImage && linkImage[1]) {
      Logger.log(`Found link image: ${linkImage[1]}`);
      return linkImage[1];
    }

    Logger.log('No image URL found');
    return null;
  } catch (error) {
    Logger.log(`Error fetching or parsing the page: ${error.toString()}`);
    return null;
  }
}

// Function to create answer slide
function createAnswerSlide(templateSlide, questionData, PREFIXES, CORRECT_ANSWER_COLOR, removeAnswerFigure, slideNumber) {
  const answerSlide = templateSlide.duplicate();
  Logger.log(`Running createAnswerSlide for question: "${questionData.question}" (Slide number: ${slideNumber})`);
  
  answerSlide.getPageElements().forEach(function(element) {
    if (element.getPageElementType() == SlidesApp.PageElementType.TEXT_BOX || element.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
      const textRange = element.asShape().getText();
      const currentText = textRange.asString();

      replaceTextInShape(element, "Year", PREFIXES.year + questionData.year);
      replaceTextInShape(element, "Question", PREFIXES.question + questionData.question);

      if (currentText.toLowerCase().includes("variant 1")) {
        replaceTextInShape(element, "Variant 1", PREFIXES.variant1 + questionData.options[0]);
        if (questionData.options[0] === questionData.correct_answer) {
          element.asShape().getFill().setSolidFill(CORRECT_ANSWER_COLOR);
        }
      } else if (currentText.toLowerCase().includes("variant 2")) {
        replaceTextInShape(element, "Variant 2", PREFIXES.variant2 + questionData.options[1]);
        if (questionData.options[1] === questionData.correct_answer) {
          element.asShape().getFill().setSolidFill(CORRECT_ANSWER_COLOR);
        }
      } else if (currentText.toLowerCase().includes("variant 3")) {
        replaceTextInShape(element, "Variant 3", PREFIXES.variant3 + questionData.options[2]);
        if (questionData.options[2] === questionData.correct_answer) {
          element.asShape().getFill().setSolidFill(CORRECT_ANSWER_COLOR);
        }
      } else if (currentText.toLowerCase().includes("variant 4")) {
        replaceTextInShape(element, "Variant 4", PREFIXES.variant4 + questionData.options[3]);
        if (questionData.options[3] === questionData.correct_answer) {
          element.asShape().getFill().setSolidFill(CORRECT_ANSWER_COLOR);
        }
      } else if (currentText.toLowerCase().includes("fact")) {
        replaceTextInShape(element, "Fact", PREFIXES.fact + questionData.fact);
      } else if (currentText.toLowerCase().includes("source")) {
        replaceTextInShape(element, "Source", PREFIXES.source + questionData.source);
      } else if (currentText.toLowerCase().includes("answer")) {
        replaceTextInShape(element, "Answer", PREFIXES.answer + questionData.correct_answer);
        if (removeAnswerFigure) {
          element.remove();
        }
      } else if (currentText.toLowerCase().includes("image")) {
        const imageUrl = getPreviewImageUrl(questionData.source);
        if (imageUrl) {
          Logger.log(`Image URL for slide number ${slideNumber}, question: "${questionData.question}": ${imageUrl}`);
          replaceShapeWithImage(answerSlide, element, imageUrl);
        }
      }
    }
  });
}

// Function to create question slide
function createQuestionSlide(templateSlide, questionData, PREFIXES, slideNumber) {
  const questionSlide = templateSlide.duplicate();
  Logger.log(`Running createQuestionSlide for question: "${questionData.question}" (Slide number: ${slideNumber})`);
  
  questionSlide.getPageElements().forEach(function(element) {
    if (element.getPageElementType() == SlidesApp.PageElementType.TEXT_BOX || element.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
      const textRange = element.asShape().getText();
      const currentText = textRange.asString();

      replaceTextInShape(element, "Year", PREFIXES.year + questionData.year);
      replaceTextInShape(element, "Question", PREFIXES.question + questionData.question);
      replaceTextInShape(element, "Variant 1", PREFIXES.variant1 + questionData.options[0]);
      replaceTextInShape(element, "Variant 2", PREFIXES.variant2 + questionData.options[1]);
      replaceTextInShape(element, "Variant 3", PREFIXES.variant3 + questionData.options[2]);
      replaceTextInShape(element, "Variant 4", PREFIXES.variant4 + questionData.options[3]);

      // Remove elements that should not be on the question slide
      if (currentText.toLowerCase().includes("fact") || currentText.toLowerCase().includes("source") || currentText.toLowerCase().includes("answer") || currentText.toLowerCase().includes("image")) {
        element.remove();
      }
    }
  });
}
