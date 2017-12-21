package com.example.hp.importfromexcel;

import android.os.Build;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.ListView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.jar.Manifest;

public class MainActivity extends AppCompatActivity {

    private static final String TAG = "MainActivity";

    private String[] FilePathString;
    private String[] FileNameString;
    private File[] listFile;
    File file;

    Button onsdCard,updir;
    ListView internalstorage;

    ArrayList<String> pathHistory;
    String lastDirectory;
    int count = 0;

    ArrayList<XYValues> uploadData;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        updir = (Button)findViewById(R.id.button1);
        onsdCard = (Button)findViewById(R.id.button2);
        internalstorage = (ListView)findViewById(R.id.list);
        uploadData = new ArrayList<>();

        CheckFilePermission();

        internalstorage.setOnItemClickListener(new AdapterView.OnItemClickListener() {
            @Override
            public void onItemClick(AdapterView<?> adapterView, View view, int i, long l) {
                lastDirectory = pathHistory.get(count);
                if(lastDirectory.equals(adapterView.getItemAtPosition(i)))
                {
                    Log.d(TAG,"InternalStorage: Selected a file for upload: "+lastDirectory);
                    readExcelData(lastDirectory);
                }
                else
                {
                    count++;
                    pathHistory.add(count,(String)adapterView.getItemAtPosition(i));
                    checkInternalStorage();
                    Log.d(TAG, "InternalStorage: "+pathHistory.get(count));
                }
            }
        });

        onsdCard.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                count =0;
                pathHistory = new ArrayList<String>();
                pathHistory.add(count,System.getenv("EXTERNAL_STORAGE"));
                Log.d(TAG, "BTNOnSDCard: "+pathHistory.get(count));
                checkInternalStorage();

            }
        });

        updir.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                if(count ==0)
                {
                    Log.d(TAG, "btnup dir: you have reached max hight..");
                }
                else
                {
                    pathHistory.remove(count);
                    count--;
                    checkInternalStorage();
                    Log.d(TAG,"btnupdir: "+pathHistory.get(count));
                }
            }
        });
    }

    private void readExcelData(String filePath)
    {
        Log.d(TAG, "ReadExccelData: Reading Excel File:");

        File inputfile = new File(filePath);

        try
        {
            InputStream inputStream = new FileInputStream(inputfile);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowsCount = sheet.getPhysicalNumberOfRows();
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            StringBuilder sb = new StringBuilder();

            for(int r=1;r<rowsCount;r++)
            {
                Row row = sheet.getRow(r);
                int cellCount = row.getPhysicalNumberOfCells();

                for(int c=0;c<cellCount;c++)
                {
                    if(c>2)
                    {
                        Log.e(TAG,"readExcelData: ERROR: Excel file format INcorrect");
                        toastMessage("ERROR: Excel File Format Is InCorrect");
                        break;
                    }
                    else
                    {
                        String value = getCellAsString(row,c,formulaEvaluator);
                        String cellInfo = "r:" +r+ ";" + "c:" +c+ ";" + "v:" +value;
                        Log.d(TAG, "ReadDataFromExcel: " +cellInfo);
                        sb.append(value +" ");
                    }
                }
                sb.append(":");
            }
            Log.d(TAG, "readExcelData: STRINGBUILDER: "+sb.toString());
            Toast.makeText(this,sb.toString(), Toast.LENGTH_SHORT).show();
            parseStringBuilder(sb);
        }
        catch(FileNotFoundException e)
        {
            Log.e(TAG, "readExcelData: FileNotFoundException: "+ e.getMessage());
        }
        catch (IOException e)
        {
            Log.e(TAG, "readExcelData: IOException: " + e.getMessage());
        }
    }

    private void parseStringBuilder(StringBuilder sb)
    {
        Log.d(TAG, "parseStringBuilder: Started parsing..");

        String[] row = sb.toString().split(":");
        for(int i=0;i<row.length;i++)
        {
            String[] columns = row[i].split(",");
            try
            {
                double x = Double.parseDouble(columns[0]);
                double y = Double.parseDouble(columns[1]);

                String cellInfo = "(x,y): ("+x+","+y+")";
                Log.d(TAG, "ParseStringBuilder: Data from row: " +cellInfo);

                uploadData.add(new XYValues(x,y));
            }
            catch(NumberFormatException e)
            {
                Log.e(TAG, "parseStringBuilder: NumberFormatException: "+e.getMessage());
            }
        }
        printDataToLog();
    }

    private void printDataToLog()
    {
        Log.d(TAG, "Printing Log DATA....");
        for(int i=0;i<uploadData.size();i++)
        {
            double x = uploadData.get(i).getX();
            double y = uploadData.get(i).getY();

            Log.d(TAG, "PrintingDataToLog: (x,y): ("+x+","+y+")");
        }
    }

    private String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator)
    {
        String value="";
        try
        {
            Cell cell = row.getCell(c);
            CellValue cellValue = formulaEvaluator.evaluate(cell);
            switch (cellValue.getCellType())
            {
                case Cell.CELL_TYPE_BOOLEAN:
                    value = ""+cellValue.getBooleanValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    double numericValue = cellValue.getNumberValue();
                    if(HSSFDateUtil.isCellDateFormatted(cell))
                    {
                        double date = cellValue.getNumberValue();
                        SimpleDateFormat formatter = new SimpleDateFormat("MM/dd/yy");
                        value = formatter.format(HSSFDateUtil.getJavaDate(date));
                    }
                    else
                    {
                        value = ""+numericValue;
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = ""+cellValue.getStringValue();
                    break;
                default:
            }
        }
        catch(NullPointerException e)
        {
            Log.e(TAG, "getCEllAsString: NullPointerException: " + e.getMessage());
        }
        return  value;
    }

    private void checkInternalStorage() {
        Log.d(TAG,"CheckInternalStorage.");
        try
        {
            if(!Environment.getExternalStorageState().equals(Environment.MEDIA_MOUNTED))
            {
                toastMessage("No SD Card Found");
            }
            else
            {
                file = new File(pathHistory.get(count));
                Log.d(TAG, "CheckExternalStorage: Directory Path: " + pathHistory.get(count));
            }

            listFile = file.listFiles();
            FilePathString = new String[listFile.length];
            FileNameString = new String[listFile.length];

            for(int i=0; i<listFile.length;i++)
            {
                FilePathString[i]=listFile[i].getAbsolutePath();
                FileNameString[i]=listFile[i].getName();
            }

            for(int i=0;i<listFile.length;i++)
            {
                Log.d("Files","FileName: "+ listFile[i].getName());
            }

            ArrayAdapter<String> adapter = new ArrayAdapter<String>(this, android.R.layout.simple_list_item_activated_1,FilePathString);
            internalstorage.setAdapter(adapter);
        }
        catch(NullPointerException e)
        {
            Log.e(TAG,"CheckInternalStorage: NULLPOINTEREXCEPTION "+e.getMessage());
        }
    }

    private void CheckFilePermission()
    {
        if(Build.VERSION.SDK_INT > Build.VERSION_CODES.LOLLIPOP)
        {
            int permissionCheck = this.checkSelfPermission("Manifest.permission.READ_EXTERNAL_STORAGE");
            permissionCheck = this.checkSelfPermission("Manifest.permission.WRITE_EXTERNAL-STORAGE");

            if(permissionCheck != 0)
            {
                this.requestPermissions(new String[]{android.Manifest.permission.READ_EXTERNAL_STORAGE, android.Manifest.permission.WRITE_EXTERNAL_STORAGE},1001);
            }
            else
            {
                Log.d(TAG , "CheckPermissions: No Need to Check Permission. SDK version < LOLLIPOP");
            }
        }
    }
    private void toastMessage(String Message)
    {
        Toast.makeText(this,Message, Toast.LENGTH_SHORT).show();
    }
}
