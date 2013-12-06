package com.example.asposeemail;

import com.aspose.email.MailAddress;
import com.aspose.email.MailMessage;
import com.aspose.email.MailMessageSaveType;

import android.os.Bundle;
import android.os.Environment;
import android.app.Activity;
import android.view.Menu;
import android.widget.TextView;

public class MainActivity extends Activity {

	public void CreateMessage()
	{
		// Create a new instance of MailMessage class
	    MailMessage message = new MailMessage();
	    // Set sender information
	    message.setFrom(new MailAddress("from@domain.com", "Sender Name", false));
	    // Add recipients
	    message.getTo().add(new MailAddress("to1@domain.com", "Recipient 1", false));
	    message.getTo().add(new MailAddress("to2@domain.com", "Recipient 2", false));
	    // Set subject of the message
	    message.setSubject("New message created by Aspose.Email for Android");
	    // Set Html body
	    message.setHtmlBody("<b>This line is in bold.</b> <br/> <br/>" +
	    "<font color=blue>This line is in blue color</font>");
	    
	    // Save message in EML, MSG and MHTML formats
	    String path = Environment.getExternalStorageDirectory().getPath();
	    message.save(path + "/Message.eml", MailMessageSaveType.getEmlFormat());
	    message.save(path + "/Message.msg", MailMessageSaveType.getOutlookMessageFormat());
	    message.save(path + "/Message.mhtml", MailMessageSaveType.getMHtmlFromat());
	}
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		final TextView tx = (TextView) findViewById(R.id.textField);

		try {
			CreateMessage();
			tx.setText("Message created successfully, please check the root of your SD card");
		} catch (Exception e) {
			tx.setText("Error during document processing: " + e.getMessage());
		}
		
	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		// Inflate the menu; this adds items to the action bar if it is present.
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}

}
