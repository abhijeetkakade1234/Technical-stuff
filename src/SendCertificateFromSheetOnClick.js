/**
 * Function to send certificates to participants listed in a Google Sheet.
 * It generates personalized certificates using a Google Slides template,
 * converts them to PDFs, and sends them via email.
 */
function sendCertificates() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues(); // Get all data from the sheet
  
    // Replace this with your Google Slides certificate template file ID
    const templateId = "Replace_with_your_template_ID";
    
    // Subject and email body to send to participants
    const subject = "Thank You for Participating in the Rangoli Competition";
    const message = `
  Dear {{name}},
  
  A heartfelt thank you for being part of our Rangoli Competition and sharing your amazing creativity! 🎨✨ Each rangoli was a burst of color and talent, adding so much beauty and joy to the event. 🌈
  
  We truly appreciate your enthusiasm and effort, which made this celebration so special. Attached is your certificate of participation 🏅—a small token of our gratitude.
  
  We can’t wait to see your talents shine again in future events! 🎉
  
  Warm regards,  
  COSA
    `;
  
    // Specify column indices (1-based) for the required data 
    // change this according to your sheet
    const nameCol = 1; // Column containing participant names
    const emailCol = 2; // Column containing participant emails
    const statusCol = 3; // Column to track certificate sending status (e.g., "Sent")
  
    // Loop through all rows in the sheet, starting from the second row (skip the header)
    for (let i = 1; i < data.length; i++) {
      const name = data[i][nameCol - 1]; // Get participant name
      const email = data[i][emailCol - 1]; // Get participant email
      const status = data[i][statusCol - 1]; // Check sending status
  
      // Skip rows with invalid or missing email addresses
      if (!email || !email.includes("@")) {
        Logger.log(`Invalid or missing email: ${email}`);
        continue;
      }
  
      // Skip rows where the certificate has already been sent
      if (status === "Sent") continue;
  
      // Generate a personalized certificate
      const certificate = createCertificate(templateId, name);
  
      // Send email with the certificate as an attachment
      MailApp.sendEmail({
        to: email.trim(),
        subject: subject,
        body: message.replace("{{name}}", name), // Replace placeholder with participant name
        attachments: [certificate.setName(`Certificate_${name}.pdf`)]
      });
  
      // Update the "Status" column to mark the certificate as sent
      sheet.getRange(i + 1, statusCol).setValue("Sent");
    }
  }
  
  /**
   * Function to create a personalized certificate for a participant.
   * It uses a Google Slides template, replaces placeholders, and converts it to a PDF.
   * @param {string} templateId - The file ID of the Google Slides template.
   * @param {string} name - The name of the participant to personalize the certificate.
   * @return {Blob} - The generated certificate as a PDF file.
   */
  function createCertificate(templateId, name) {
    try {
      Logger.log(`Attempting to access template with ID: ${templateId}`);
      const template = DriveApp.getFileById(templateId); // Access the template file
      Logger.log("Successfully accessed the template.");
  
      // Create a copy of the template with a unique name
      const copy = template.makeCopy(`Certificate - ${name}`);
      const slideDeck = SlidesApp.openById(copy.getId());
  
      // Replace placeholder text in all slides with the participant's name
      const slides = slideDeck.getSlides();
      slides.forEach(slide => {
        slide.replaceAllText('{{Name}}', name);
      });
  
      // Save and close the modified slide
      slideDeck.saveAndClose();
  
      // Convert the modified slide to a PDF
      const pdf = DriveApp.getFileById(copy.getId()).getAs(MimeType.PDF);
  
      // Clean up: Move the slide copy to the trash
      DriveApp.getFileById(copy.getId()).setTrashed(true);
  
      return pdf; // Return the generated PDF
    } catch (error) {
      Logger.log(`Error accessing file: ${error.message}`);
      throw error; // Re-throw the error for debugging purposes
    }
  }
  