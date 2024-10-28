//This code is to automate certificate creation as soon as google form is submitted
// paste this code in trigger 
function onFormSubmit(e) {
  // Check if the event object is defined
  if (!e || !e.values) {
    Logger.log("Event object is undefined or doesn't contain values");
    return;
  }

  const responses = e.values;
  const email = responses[2]; // Adjust index based on form structure
  const name = responses[1]; // Adjust index based on form structure

  // Google Slide template ID
  const slideTemplateId = 'Replace_with_your_template'; // Replace with your template ID
  const folderId = 'Replace_with_your_Drive_folder_ID'; // Replace with your Drive folder ID for saving certificates

  // Create a copy of the template in the specified folder
  const slideTemplate = DriveApp.getFileById(slideTemplateId).makeCopy(`Certificate for ${name}`, DriveApp.getFolderById(folderId));
  const slide = SlidesApp.openById(slideTemplate.getId());
  const slidePage = slide.getSlides()[0];

  // Replace placeholder {{Name}} with the actual name
  slidePage.replaceAllText('{{Name}}', name);
  // certificate should contain a place holder ->> {{Name}}  This is mandatory!!!!!!!!!!!!!!!!!!!!!!!!!

  // Refresh the slide to ensure changes are processed
  slide.saveAndClose();

  // Re-open the slide to ensure the changes are saved
  const updatedSlide = DriveApp.getFileById(slideTemplate.getId()).getAs(MimeType.PDF);

  // Send email with the certificate attached
  MailApp.sendEmail({
    to: email,
    subject: "Your Participation Certificate",
    body: `Dear ${name},\n\nThank you for participating! Attached is your participation certificate.\n\nBest regards,\nEvent Team`,
    attachments: [updatedSlide.setName(`Certificate_${name}.pdf`)]
  });

  // Cleanup: Move the generated certificate PDF to the specified folder and trash the slide copy
  DriveApp.getFileById(slideTemplate.getId()).setTrashed(true);
}
