
package com.example.asposeemailpst;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

import com.aspose.email.FileFormatVersion;
import com.aspose.email.FolderInfo;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiContactElectronicAddress;
import com.aspose.email.MapiContactGender;
import com.aspose.email.MapiContactNamePropertySet;
import com.aspose.email.MapiContactProfessionalPropertySet;
import com.aspose.email.MapiContactTelephonePropertySet;
import com.aspose.email.MapiJournal;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiNote;
import com.aspose.email.MapiRecipientCollection;
import com.aspose.email.MapiRecipientType;
import com.aspose.email.MapiTask;
import com.aspose.email.MapiTaskHistory;
import com.aspose.email.MapiTaskOwnership;
import com.aspose.email.NoteColor;
import com.aspose.email.PersonalStorage;
import com.aspose.email.StandardIpmFolder;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	private static PersonalStorage pst = null;
	
	private static String strBaseFolder = Environment.getExternalStorageDirectory().getPath() + "/";
	
	public static void RunPstExample()
	{		
		//Create an instance of PersonalStorage
		pst = PersonalStorage.create(strBaseFolder + "sample.pst", 0);

		//Add Messages to PST
		addMessagesToPst();
		
		//Add MapiTask to PST
		addMapiTaskToPst();
		
		//Add Contacts To PST
		addContactsToPst();
		
		//Add Journal To PST
		addJournalToPst();
		
		//Add MapiCalendar To PST
		addMapiCalendarToPst();
		
		//Add MapiNote To PST
		addMapiNoteToPst();
		
		pst.dispose();
	}
	
	public static void addMessagesToPst()
	{
		//Create a folder at root of pst
		pst.getRootFolder().addSubFolder("myInbox");		

		//Add message to newly created folder
		pst.getRootFolder().getSubFolder("myInbox").addMessage(MapiMessage.fromFile(strBaseFolder + "New-out.msg"));
	}
	
	public static void addMapiTaskToPst()
	{
		MapiTask task = new MapiTask("To Do", "Just click and type to add new task", new Date(), new Date());
		task.setPercentComplete(20);
		task.setEstimatedEffort(2000);
		task.setActualEffort(20);
		task.setHistory(MapiTaskHistory.Assigned);
		task.setLastUpdate(new Date());
		task.getUsers().setOwner("Darius");
		task.getUsers().setLastAssigner("Harkness");
		task.getUsers().setLastDelegate("Harkness");
		task.getUsers().setOwnership(MapiTaskOwnership.AssignersCopy);
				
		FolderInfo taskFolder = pst.createPredefinedFolder("Tasks", StandardIpmFolder.Tasks);
		taskFolder.addMapiMessageItem(task);
		
		pst.dispose();
	}

	public static void addContactsToPst()
	{
		// Contact #1
		MapiContact contact1 = new MapiContact("Sebastian Wright", "SebastianWright@dayrep.com");

		// Contact #2
		MapiContact contact2 = new MapiContact("Wichert Kroos", "WichertKroos@teleworm.us", "Grade A Investment");

		// Contact #3
		MapiContact contact3 = new MapiContact("Christoffer van de Meeberg", "ChristoffervandeMeeberg@teleworm.us", "Krauses Sofa Factory", "046-630-4614");

		// Contact #4
		MapiContact contact4 = new MapiContact();
		contact4.setNameInfo(new MapiContactNamePropertySet("Margaret", "J.", "Tolle"));
		contact4.getPersonalInfo().setGender(MapiContactGender.Female);
		contact4.setProfessionalInfo(new MapiContactProfessionalPropertySet("Adaptaz", "Recording engineer"));
		contact4.getPhysicalAddresses().getWorkAddress().setAddress("4 Darwinia Loop EIGHTY MILE BEACH WA 6725");
		contact4.getElectronicAddresses().setEmail1(new MapiContactElectronicAddress("Hisen1988", "SMTP", "MargaretJTolle@dayrep.com"));
		contact4.getTelephones().setBusinessTelephoneNumber("(08)9080-1183");
		contact4.getTelephones().setMobileTelephoneNumber("(925)599-3355");

		// Contact #5
		MapiContact contact5 = new MapiContact();
		contact5.setNameInfo(new MapiContactNamePropertySet("Matthew", "R.", "Wilcox"));
		contact5.getPersonalInfo().setGender(MapiContactGender.Male);
		contact5.setProfessionalInfo(new MapiContactProfessionalPropertySet("Briazz", "Psychiatric aide"));
		contact5.getPhysicalAddresses().getWorkAddress().setAddress("Horner Strasse 12 4421 SAASS");
		contact5.getTelephones().setBusinessTelephoneNumber("0650 675 73 30");
		contact5.getTelephones().setHomeTelephoneNumber("(661)387-5382");

		// Contact #6
		MapiContact contact6 = new MapiContact();
		contact6.setNameInfo(new MapiContactNamePropertySet("Bertha", "A.", "Buell"));
		contact6.setProfessionalInfo(new MapiContactProfessionalPropertySet("Awthentikz", "Social work assistant"));
		contact6.getPersonalInfo().setPersonalHomePage("B2BTies.com");
		contact6.getPhysicalAddresses().getWorkAddress().setAddress("Im Astenfeld 59 8580 EDELSCHROTT");
		contact6.getElectronicAddresses().setEmail1(new MapiContactElectronicAddress("Experwas", "SMTP", "BerthaABuell@armyspy.com"));
		contact6.setTelephones(new MapiContactTelephonePropertySet("06605045265"));

		FolderInfo contactFolder = pst.createPredefinedFolder("Contacts", StandardIpmFolder.Contacts);
		contactFolder.addMapiMessageItem(contact1);
		contactFolder.addMapiMessageItem(contact2);
		contactFolder.addMapiMessageItem(contact3);
		contactFolder.addMapiMessageItem(contact4);
		contactFolder.addMapiMessageItem(contact5);
		contactFolder.addMapiMessageItem(contact6);

	}
	
	public static void addJournalToPst()
	{
		Date d1 = new Date();
		Calendar cl = Calendar.getInstance();
		cl.setTime(d1);
		cl.add(Calendar.HOUR, 1);
		Date d2 = cl.getTime();

		MapiJournal journal = new MapiJournal("daily record", "called out in the dark", "Phone call", "Phone call");
		journal.setStartTime(d1);
		journal.setEndTime(d2);		
		FolderInfo journalFolder = pst.createPredefinedFolder("Journal", StandardIpmFolder.Journal);
		journalFolder.addMapiMessageItem(journal);
		
	}
	
	public static void addMapiCalendarToPst()
	{
		// Create the appointment
		Calendar calendar = Calendar.getInstance(TimeZone.getTimeZone("GMT"));
		calendar.set(2012, Calendar.JUNE, 23, 0, 0, 0);
		Date startDate = calendar.getTime();
		calendar.set(2012, Calendar.JUNE, 23);
		Date endDate = calendar.getTime();

		MapiCalendar appointment = new MapiCalendar(
		 "LAKE ARGYLE WA 6743",
		 "Appointment",
		 "This is a very important meeting :)",
		 startDate,
		 endDate);

		// Create the meeting
		MapiRecipientCollection attendees = new MapiRecipientCollection();
		attendees.add("ReneeAJones@armyspy.com", "Renee A. Jones", MapiRecipientType.MAPI_TO);
		attendees.add("SzllsyLiza@dayrep.com", "Szollosy Liza", MapiRecipientType.MAPI_TO);

		MapiCalendar meeting = new MapiCalendar(
		 "Meeting Room 3 at Office Headquarters",
		 "Meeting",
		 "Please confirm your availability.",
		 startDate,
		 endDate,
		 "CharlieKhan@dayrep.com",
		 attendees
		 );
		
		FolderInfo calendarFolder = pst.createPredefinedFolder("Calendar", StandardIpmFolder.Appointments);
		calendarFolder.addMapiMessageItem(appointment);
		calendarFolder.addMapiMessageItem(meeting);
	}
	
	public static void addMapiNoteToPst()
	{
		//Load the template MapiNote
		MapiMessage mess = MapiMessage.fromFile(strBaseFolder + "/mnt/sdcard/Note.msg");

		// Note #1
        MapiNote note1 = (MapiNote)mess.toMapiMessageItem();
        note1.setSubject("Yellow color note");
        note1.setBody("This is a yellow color note");

        // Note #2
        MapiNote note2 = (MapiNote)mess.toMapiMessageItem();
        note2.setSubject("Pink color note");
        note2.setBody("This is a pink color note");
        note2.setColor(NoteColor.Pink);

        // Note #3
        MapiNote note3 = (MapiNote)mess.toMapiMessageItem();
        note2.setSubject("Blue color note");
        note2.setBody("This is a blue color note");
        note2.setColor(NoteColor.Blue);
        note3.setHeight(500);
        note3.setWidth(500);

        PersonalStorage pst = PersonalStorage.create(strBaseFolder + "MapiNoteToPST.pst", FileFormatVersion.Unicode);
        FolderInfo notesFolder = pst.createPredefinedFolder("Notes", StandardIpmFolder.Notes);
        notesFolder.addMapiMessageItem(note1);
        notesFolder.addMapiMessageItem(note2);
        notesFolder.addMapiMessageItem(note3);		
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		//final TextView tx = (TextView) findViewById(R.id.textField1);
		RunPstExample();
		//tx.setText("Message Created successfully");
	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}

}
