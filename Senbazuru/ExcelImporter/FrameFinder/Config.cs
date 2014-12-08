using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameFinder
{
    class Config
    {
        // this could be modified to absolute path
        public static String BASEDIR = "../../../resources/FrameFinder";
        // directory to store the original spreadsheets
        public static String SHEETDIR = BASEDIR + "/data/sheets";
        // directory to store the output:
        // each spreadsheet labeled with semantic labels for each row
        public static String OUTPUTDIR = BASEDIR + "/data/results";

        // files to store intermediate results
        public static String CRFTEMPDIR = BASEDIR + "/data/tmp";
        public static String CRFTMPFEATURE = CRFTEMPDIR + "/feature.tmp";
        public static String CRFTMPPREDICT = CRFTEMPDIR + "/predict.tmp";

        // template file for CRF++ to parse the provided features
        public static String CRFPPTEMPLATEPATH = BASEDIR + "/data/template";
        // training data
        public static String CRFTRAINDATAPATH = BASEDIR + "/data/saus_train.data";

        /*****************************************
        * please specify the directory of CRF++
        *****************************************/
        // directory of installed CRF++
        public static String CRFPPDIR = BASEDIR + "/CRF++-0.58";
    }
}
