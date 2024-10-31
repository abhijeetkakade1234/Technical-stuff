// This code is to automate certificate creation as soon as a Google Form is submitted
// To use: paste this code in the Apps Script editor and set up a trigger
// Go to Extensions > Apps Script, paste code, and set up an "On form submit" trigger

function onFormSubmit(e) {
  // Check if the event object is defined to ensure valid form submission data
  if (!e || !e.values) {
    Logger.log("Event object is undefined or doesn't contain values"); // Log an error if event data is missing
    return; // Stop execution if no event data is found
  }

  // Capture form responses from the event object
  const responses = e.values;
  const email = responses[2]; // Fetch email from form responses (adjust index based on form structure)
  const name = responses[1]; // Fetch participant's name from form responses (adjust index based on form structure)

  // Define Google Slide template ID (where {{Name}} placeholder is located) This is FUCKING MANDATORY !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  const slideTemplateId = 'Replace_with_your_template'; // Replace with your Google Slide template ID here
  const folderId = 'Replace_with_your_Drive_folder_ID'; // Replace with your Google Drive folder ID to save certificates

  // Create a copy of the certificate template for each submission and store it in the specified Drive folder
  const slideTemplate = DriveApp.getFileById(slideTemplateId).makeCopy(`Certificate for ${name}`, DriveApp.getFolderById(folderId));
  const slide = SlidesApp.openById(slideTemplate.getId()); // Open the copied slide as a Google Slides document
  const slidePage = slide.getSlides()[0]; // Access the first slide page in the document

  // Replace placeholder text {{Name}} in the template with the actual participant's name
  slidePage.replaceAllText('{{Name}}', name);
  // Note: Ensure the slide template contains a placeholder {{Name}} for this replacement to work

  // Save and close the modified slide to apply changes
  slide.saveAndClose();

  // Convert the modified slide to PDF format for emailing
  const updatedSlide = DriveApp.getFileById(slideTemplate.getId()).getAs(MimeType.PDF);

  // Send an email with the certificate PDF as an attachment
  MailApp.sendEmail({
    to: email, // Email address of the participant
    subject: "Your Participation Certificate", // Subject of the email
    body: `Dear ${name},\n\nThank you for participating! Attached is your participation certificate.\n\nBest regards,\nEvent Team`,
    attachments: [updatedSlide.setName(`Certificate_${name}.pdf`)] // Attach the PDF certificate
  });

  // Clean up: Move the generated slide copy to trash after the PDF is created
  DriveApp.getFileById(slideTemplate.getId()).setTrashed(true);
  // IF u want to restore the certificate u can find it in the trash folder in the admin drive
}
