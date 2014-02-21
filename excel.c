#include<stdio.h>
#include<string.h>
#include<memory.h>
#include<stdint.h>


#define STDLEN 32

#define boolean int
#define False 0
#define True  1

#define MAXSHEET 255   
#define MAXSTYLE 255

#define DE_WIDTH 80
#define DE_HEIGHT 20


/*



*/

struct Cell{
	void *data;
	short mergeDown;
	short mergeAcross;
};


struct ExcelSheet{
    char title[STDLEN+1];   //sheet Title
    int row;                //sheet rows number
    int col;                //sheet column number
    int defaultWidth;       //the default column width
    int defaultHeight;      //the default  row height
    int* colwidth;          //the array of column width 
    int* rowheight;          //the array of row height
    int32_t * stylemap;         // the bit map of style 
    void* data[];           // the  array of cells's data 
};


struct ExcelStyle{
    int id;
    char name[STDLEN+1];
    int valign;             //Vertical align
    int halign;             //  h       align
    char font_name[STDLEN]; //  font name
    int font_size;          //font size
    int isbold;             //bold font
};

struct Excel{
    char author[STDLEN+1];     
    char lastAuthor[STDLEN+1]; 
    char created[20];           
    char company[STDLEN+1];   
    double version;             

    long    WindowHeight;       //default 6795
    long    WindowWidth;        //default 8460
    long    WindowTopX;         //default 120
    long    WindowTopY;         //default 15
    boolean    ProtectStructure;        //default:false
    boolean     ProtectWindows;         //default:false
    struct ExcelSheet* sheets[MAXSHEET];
    int sheetnum;
    struct ExcelStyle* styles[MAXSTYLE];
    int stylenum;
};


/*
        filedir      ÎÄ¼þÂ·¾¶
        sheetnum     SheetÊýÁ¿
        stylenum     StyleÊýÁ¿
*/
struct Excel * new_excel(){
    struct Excel *tmp=(struct Excel *)malloc(sizeof(struct Excel));
    tmp->sheetnum=0;
    tmp->stylenum=0;
    memset(tmp,'\0',sizeof(struct Excel));
    return tmp;
}


struct ExcelStyle * new_style(int id,char * name){
    struct ExcelStyle *tmp=(struct ExcelStyle *)malloc(sizeof(struct ExcelStyle));
    tmp->id=id;
    strncpy(tmp->name,name,STDLEN);
    return tmp;
}

void Style_setValign(struct ExcelStyle * tmp){

}

void Style_setBold(struct ExcelStyle * tmp,boolean isbold){
    tmp->isbold=isbold;
}

void Style_setFontSize(struct ExcelStyle * tmp,int size){
    tmp->font_size=size;
}


struct ExcelSheet * new_sheet(char *title,int row,int col){
    struct ExcelSheet *tmp;
    int size=sizeof(struct ExcelSheet)+row*col*(sizeof(void*));
    tmp=(struct ExcelSheet *)malloc(size);
    memset(tmp,'\0',size);
    strncpy(tmp->title,title,STDLEN);
    tmp->row=row;
    tmp->col=col;
    //  alloc colwidth array
    tmp->colwidth=malloc(sizeof(int)*col);
    memset(tmp->colwidth,'\0',sizeof(int)*col);
    tmp->defaultWidth=DE_WIDTH;

    for(int i=0;i<col;i++){
        tmp->colwidth[i]=0;  //
    }

      //  alloc colwidth array
    tmp->rowheight=malloc(sizeof(int)*row);
    memset(tmp->rowheight,'\0',sizeof(int)*row);
    tmp->defaultHeight=DE_HEIGHT;

      for(int i=0;i<row;i++){
        tmp->rowheight[i]=0;  //
    }

    //
    tmp->stylemap=malloc(sizeof(int32_t)*row*col);
   

    return tmp;
}

void add_sheet_to_excel(struct Excel * xls,struct ExcelSheet * sht){
    int curidx=xls->sheetnum;
    xls->sheets[curidx]=sht;
    xls->sheetnum++;

}

void add_style_to_excel(struct Excel * xls,struct ExcelStyle * sty){
    int curidx=xls->stylenum;
    xls->styles[curidx]=sty;
    xls->stylenum++;

}

void dumpSht(struct ExcelSheet * sht){
    printf("row:%d\n",sht->row);
    printf("col:%d\n",sht->col);

    for(int i=0;i<sht->row;i++)
            for(int j=0;j<sht->col;j++){
                int lo=i*sht->col+j;
                if(sht->data[lo]!=NULL){
                    printf("%d:%d=>%s\n",i,j,sht->data[lo]);
                }
            }
}


void dump(struct Excel *xls){
    printf("sheetnum:%d\n",xls->sheetnum);


    for(int i=0;i<xls->sheetnum;i++){

            dumpSht(xls->sheets[i]);
    }
    

}

static char *xmlHeadStr="<?xml version=\"1.0\"?>\n<?mso-application progid=\"Excel.Sheet\"?>";





char * Excel2Xml(struct Excel * est,char * filename){
    FILE * fp;
    fp=fopen("1.xls","w");
    if(fp==NULL){
        printf("open file fail");
    }

    fputs(xmlHeadStr,fp);
    fputc('\n',fp);
    fprintf(fp,"<ss:Workbook xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
    fputc('\n',fp);
   // fprintf(fp,"<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\"</DocumentProperties>");
   // fputc('\n',fp);

    for(int i=0;i<est->sheetnum;i++){
            struct ExcelSheet * tmp=est->sheets[i];

            fprintf(fp,"<ss:Worksheet ss:Name=\"%s\"> <ss:Table>",tmp->title); 
            fputc('\n',fp);

            //set the column width
            for(int i=0;i<tmp->col;i++){
                  fprintf(fp,"<ss:Column ss:Width=\"%d\"/>",getColWidth(tmp,i));
                  fputc('\n',fp);
            }
          


            for(int j=0;j<tmp->row;j++){
                        fprintf(fp,"<ss:Row ss:Height=\"%d\">",getRowHeight(tmp,j));
                        fputc('\n',fp);
                        for(int k=0;k<tmp->col;k++){
                            int curidx=tmp->col*j+k;
                            fprintf(fp,"<ss:Cell>");
                            fputc('\n',fp);
                            if(tmp->data[curidx]!=NULL)
                                 fprintf(fp," <ss:Data ss:Type=\"String\">%s</ss:Data>",tmp->data[curidx]);
                             else 
                                 fprintf(fp," <ss:Data ss:Type=\"String\"></ss:Data>");
                            fputc('\n',fp);
                            fprintf(fp,"</ss:Cell>");
                            fputc('\n',fp);
                        }
                        fprintf(fp,"</ss:Row>");
                        fputc('\n',fp);
            
            }
            fprintf(fp,"</ss:Table>\n</ss:Worksheet>");
            fputc('\n',fp);
            
                
    }

    fputs("</Workbook>",fp);
    fputc('\n',fp);
    fclose(fp);
}



/*
    set Cell value

*/
int setCell(struct ExcelSheet * xls,int y,int x,char * str){
    int index=y*(xls->col)+x;
    int length=strlen(str)+1;
    xls->data[index]=malloc(length);
    memset(xls->data[index],'\0',length);
    strncpy(xls->data[index],str,length);
}

/*
         set Col Width
*/
int setColWidth(struct ExcelSheet * sht,int idx,int width){
    if(idx>=sht->col)   return False;
    if(idx<0)   return False;
    sht->colwidth[idx]=width;
    return True;
}

/*
         set Row height
*/
int setRowHeight(struct ExcelSheet * sht,int idx,int height){
    if(idx>=sht->row)   return False;
    if(idx<0)   return False;
    sht->rowheight[idx]=height;
    return True;
}


/*
    set the default column width;
*/
int setDefaultColWidth(struct ExcelSheet * sht,int width){
    sht->defaultWidth=width;
}

/*
    set the default column width;
*/
int setDefaultRowHeight(struct ExcelSheet * sht,int height){
    sht->defaultHeight=height;
}
/*
    return col width;
    if the colwidth is not set return defaultWidth
*/
int getColWidth(struct ExcelSheet * sht,int idx){
    int colw=sht->colwidth[idx];
    if(colw!=0)     return colw;        //if the col width is not null, return it 
    colw=sht->defaultWidth;
    if(colw!=0)     return colw;        
    return DE_WIDTH;

}

/*
    return row height;
    if the row height is not set return defaultHeight
*/
int getRowHeight(struct ExcelSheet * sht,int idx){
    printf("%d start\n",idx);
    int colw=sht->rowheight[idx];

    if(colw!=0)     return colw;        //if the col width is not null, return it 
    colw=sht->defaultHeight;
    if(colw!=0)     return colw;        
    return DE_HEIGHT;
    

}



int main(){

    printf("%d|%d|%d|%d\n",sizeof(int16_t),sizeof(short),sizeof(char),sizeof(void));

    struct ExcelStyle *sty1=new_style(1,"st1");
    Style_setBold(sty1,True);


    struct Excel *xls=new_excel();
    struct ExcelSheet *sht1=new_sheet("1",5,6); 
    setCell(sht1,1,1,"haha11");
    setCell(sht1,1,2,"haha12");
    setCell(sht1,4,5,"haha45");
    setCell(sht1,0,5,"haha05");
    setDefaultColWidth(sht1,100);
    setColWidth(sht1,0,100);
    setColWidth(sht1,1,120);
    setColWidth(sht1,2,140);
    setColWidth(sht1,3,160);
    setColWidth(sht1,4,180);
    add_sheet_to_excel(xls,sht1);

    struct ExcelSheet * sht2=new_sheet("2",7,6); 
    setCell(sht2,1,1,"2haha11");
    setCell(sht2,1,2,"2haha12");
    setCell(sht2,4,5,"2haha45");
    setCell(sht2,0,5,"2haha05");
    setRowHeight(sht2,1,70);
    add_sheet_to_excel(xls,sht2);
    Excel2Xml(xls,"1.xls");
    dump(xls);
    return 0;
}


