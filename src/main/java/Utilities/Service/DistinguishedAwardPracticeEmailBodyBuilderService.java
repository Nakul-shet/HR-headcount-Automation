package Utilities.Service;

import Utilities.Common.EmailBodyUtilities;
import Utilities.Common.ExcelUtilities;
import Utilities.Configuration.SpotAwardConfig;

public class DistinguishedAwardPracticeEmailBodyBuilderService {
    public static String buildEligibilityEmailBody() throws Exception {

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");

        htmlBuilder.append("<p>Dear All,</p>");

        htmlBuilder.append("<p>Please share the R&R nominations from your respective teams for Quarter ")
                .append(EmailBodyUtilities.getCurrentQuarter())
                .append(", ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append("<p>Kindly share the nominations as per the <b>New Org</b> structure by clicking the <b>Nominate Employees</b> button below, on or before 28 ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append("<ul style='font-family:Arial,sans-serif; font-size:14px; color:#333;'>")
                .append("<li><strong>Distinguished Performer of the Quarter</strong> ")
                .append("(Individuals who have demonstrated outstanding dedication, achievements, and contributed significantly to TGS success).</li>")
                .append("</ul>");

        htmlBuilder.append(EmailBodyUtilities.nominateEmployeesButton(SpotAwardConfig.ORG_PRACTICE));

        htmlBuilder.append("<p><b>Distinguished Performer Award Eligibility</b></p>");

        htmlBuilder.append("<table border='1' style='border-collapse: collapse; width: auto; border-width: 2px; text-align:center;'>");
        htmlBuilder.append("<tr style='background-color:yellow'>");
        htmlBuilder.append("<th style='padding: 2px 30px; border: 2px solid black;'>New Org</th>");
        htmlBuilder.append("<th style='padding: 2px 30px; border: 2px solid black;'>")
                .append(" Eligibility</th>");
        htmlBuilder.append("</tr>");

        htmlBuilder.append(ExcelUtilities.readHeadCountData("Practice_Distinguished"));
        htmlBuilder.append("</table>");

        htmlBuilder.append("<ul style='font-family:Arial,sans-serif; font-size:14px; color:#333;'>")
                .append("<p><strong>Please Note:</strong></p>")
                .append("<li>The citations provided for award winners will be shared to all employees during the award announcement.</li>")
                .append("<li>Request you to add details like why this award is bestowed to the individual, highlight the achievements and contributions in Q")
                .append(EmailBodyUtilities.getCurrentQuarter())
                .append(".</li>")
                .append("<li>Two liner citations will not be accepted. Word limit: 100–150 words.</li>")
                .append("</ul>");

        htmlBuilder.append(EmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</body></html>");
        return htmlBuilder.toString();
    }

    public static String buildReminderEmailBody1() throws Exception {

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");
        htmlBuilder.append("<p>Dear Managers,</p>");

        htmlBuilder.append("<p>This is your <b style='color : purple;'>gentle-but-not-so-gentle reminder</b> to submit those <b style='color : purple;'>Distinguished Award Nominations</b> before 28 ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append("<p>Don’t let those <b style='color : purple;'>silent rockstars go unnoticed</b>. It’s your chance to shine a light on your team’s awesomeness 🌟.</p>");

        htmlBuilder.append("<p>If you've already submitted, <b style='color : purple;'> you’re officially awesome </b> and can proudly ignore this message.</p>");

        htmlBuilder.append("<p>If not – tick tock ⏰… recognition season is calling!</p>");

        htmlBuilder.append("<p>Got questions or need help? Ping the ever-helpful folks at ")
                .append("<b><a href='mailto:TGSHRIndiaOps@teksystems.com'>TGSHRIndiaOps@teksystems.com</a></b>")
                .append("</p>")
                .append("<p>Let the nominations roll in!</p>");

        htmlBuilder.append(EmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</body></html>");

        return htmlBuilder.toString();

    }

    public static String buildReminderEmailBody2() throws Exception {

        StringBuilder htmlBuilder = new StringBuilder();

        htmlBuilder.append("<html><body>");
        htmlBuilder.append("<p>Dear All,</p>");

        htmlBuilder.append("<p>This is your <b>final, no-kidding, last-chance, curtain-call reminder</b> to submit your <b>Distinguished Award Nominations</b> before 28 ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append("<p>The nomination window <b>slams shut</b> at the end of the day on 28 ")
                .append(java.time.LocalDate.now().getMonth().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.ENGLISH))
                .append(" ")
                .append(java.time.LocalDate.now().getYear())
                .append("</p>");

        htmlBuilder.append("After that, even puppy eyes or “I forgot” won’t work. 😅</p>");

        htmlBuilder.append("<p><b>Still haven’t nominated?</b><br>");
        htmlBuilder.append("Please don’t be that manager whose team says, “Recognition? Never heard of it.”</p>");

        htmlBuilder.append("<p>Let’s not disappoint the unsung heroes quietly saving the day in your team!</p>");

        htmlBuilder.append("<p><b>Already submitted?</b> You’re a legend – please ignore this email and go treat yourself to a coffee. ☕</p>");

        htmlBuilder.append("<p>For any last-minute confusion or friendly SOS, reach out to the award-wielding champs at:<br>");
        htmlBuilder.append("📧 <b><a href='mailto:TGSHRIndiaOps@teksystems.com'>TGSHRIndiaOps@teksystems.com</a></b></p>");

        htmlBuilder.append("<p><b>Let’s make those nominations count</b> (before the HR ops team starts chasing you with memes)! 😄</p>");

        htmlBuilder.append(EmailBodyUtilities.getEmailSignature());
        htmlBuilder.append("</body></html>");

        return htmlBuilder.toString();

    }
}
