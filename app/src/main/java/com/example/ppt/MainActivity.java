package com.example.ppt;

import androidx.appcompat.app.AppCompatActivity;

import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.os.Bundle;
import android.text.method.ScrollingMovementMethod;
import android.util.Log;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.xslf.extractor.XSLFPowerPointExtractor;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.w3c.dom.Text;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
//import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.util.Removal;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideShow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.xmlbeans.XmlException;


public class MainActivity extends AppCompatActivity {

    public static InputStream InputstreamGeneralPhotos1;
    public static InputStream InputstreamGeneralPhotos2;
    public static InputStream InputstreamGeneralPhotos3;
    public static InputStream InputstreamGeneralPhotos4;
    public static InputStream InputstreamGeneralPhotos5;
    public static InputStream InputstreamGeneralPhotos6;
    public static InputStream InputstreamBicyclePhotos1;
    public static InputStream InputstreamBicyclePhotos2;
    public static InputStream InputstreamFourhoursPhotos1;
    public static InputStream InputstreamFourhoursPhotos2;
    public static InputStream InputstreamFourhoursPhotos3;
    public static InputStream InputstreamFourhoursPhotos4;

    static {
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLInputFactory",
                "com.fasterxml.aalto.stax.InputFactoryImpl"
        );
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLOutputFactory",
                "com.fasterxml.aalto.stax.OutputFactoryImpl"
        );
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLEventFactory",
                "com.fasterxml.aalto.stax.EventFactoryImpl"
        );
    }

    TextView textView;

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        MenuInflater inflater = getMenuInflater();
        inflater.inflate(R.menu.menu,menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        switch(item.getItemId()){
            case R.id.item:
                openDocumentFromFileManager();
                return true;
//            case R.id.item1:
//                return true;
            default:
                return super.onOptionsItemSelected(item);
        }
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        textView = (TextView)findViewById(R.id.textView);
    }

    private void openDocumentFromFileManager() {
        //this is the action to open doc file from file manager
        Intent i = new Intent();
        i.setType("application/*");
        i.setAction(Intent.ACTION_GET_CONTENT);
        if (PermissionsHelper.getPermission(this, android.Manifest.permission.WRITE_EXTERNAL_STORAGE, R.string.title_storage_permission
                , R.string.text_storage_permission, 1111)) {
            startActivityForResult(Intent.createChooser(i, "Select Document"), 111);
        }
    }


    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        try {
            if (resultCode==RESULT_OK){
                switch (requestCode){
                    case 111:
                        //this is action performed after openDocumentFromFileManager() when doc is selected
                        FileInputStream inputStream = (FileInputStream) getContentResolver().openInputStream(data.getData());
                        XMLSlideShow ppt = new XMLSlideShow(inputStream);

                        extractImages(ppt);

                        int slides = ppt.getSlides().size();
                        Log.i("Slides",Integer.toString(slides));


//                        extractText(docx);

//                        List<XWPFParagraph> paragraphList = docx.getParagraphs();
//
//                        for (XWPFParagraph paragraph:paragraphList){
//                            String paragrapthText = paragraph.getText();
//                            System.out.println(paragrapthText);
//                        }


                }
            }else {
                Toast.makeText(this, "Your file is not loaded", Toast.LENGTH_SHORT).show();
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void extractImages(XMLSlideShow ppt){
        try {

            //function to retrieve all the images from the doc file that have been selected
//
//            HWPFDocument wordDoc = new HWPFDocument(new FileInputStream(fileName));
//
//            XWPFWordExtractor extractor = new XWPFWordExtractor(docx);
//
////            extractText(extractor);
//
//            System.out.println(extractor.getText());

            XSLFPowerPointExtractor extractor = new XSLFPowerPointExtractor(ppt);

            textView.setText(extractor.getText());
            textView.setMovementMethod(new ScrollingMovementMethod());



//            List<XWPFPictureData> picList = docx.getAllPackagePictures();
//
//            int i = 1;
//
//            for (XWPFPictureData pic : picList) {
//
//                System.out.print(pic.getPictureType());
//                System.out.print(pic.getData());
//                System.out.println("Image Number: " + i + " " + pic.getFileName());
//                switch (i){
//                    case 1:
//                        InputstreamGeneralPhotos1 = new ByteArrayInputStream(pic.getData());
//                        if (InputstreamGeneralPhotos1 != null){
//                            Bitmap selectedImage = BitmapFactory.decodeStream(InputstreamGeneralPhotos1);
//                            img_generalPhotos1.setImageBitmap(selectedImage);
////                            generalPhotosUnSelect1.setVisibility(View.VISIBLE);
//                        }
//                        break;
//                    case 2:
//                        InputstreamGeneralPhotos2 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 3:
//                        InputstreamGeneralPhotos3 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 4:
//                        InputstreamGeneralPhotos4 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 5:
//                        InputstreamGeneralPhotos5 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 6:
//                        InputstreamGeneralPhotos6 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 7:
//                        InputstreamBicyclePhotos1 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 8:
//                        InputstreamBicyclePhotos2 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 9:
//                        InputstreamFourhoursPhotos1 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 10:
//                        InputstreamFourhoursPhotos2 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 11:
//                        InputstreamFourhoursPhotos3 = new ByteArrayInputStream(pic.getData());
//                        break;
//                    case 12:
//                        InputstreamFourhoursPhotos4 = new ByteArrayInputStream(pic.getData());
//                        break;
//                }
//                i++;
//            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }
}
