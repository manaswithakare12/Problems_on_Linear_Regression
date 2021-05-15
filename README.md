# BetterTeaching
## _Personalized Teaching Feedback_

[![Build Status](https://travis-ci.org/joemccann/dillinger.svg?branch=master)](https://travis-ci.org/joemccann/dillinger)

<center><img src="./templates/audioMonoPlot.png" width="900" height="200"></center>

[Sample report](https://github.com/UnpredictablePrashant/BetterTeaching/blob/main/out/output.pdf)

### Introduction
Teaching is a noble profession. A good teacher is considered to have various skills from class management to communication. This project deals with identifying the gap on communication of teachers/speakers and provide them with a detailed personalized feedback where they can improve.

The "BetterTeaching" is an AI powered personal coach which focuses on helping individuals become better in communication especially required for becoming good speaker or good teacher.

### Factors effecting the report
The project is divided into two main part, in the first part the speaking/teaching video is compared using the static analysis which include the `text analysis`, `voice analysis`, `facial & gesture analysis`. While in the second part, the analysis is done using the CNN model based on the features of top speakers/educators. The entire module is completely AI powered.

Factors are:
- ✅Average Length of sentence
- ✅Word repetition
- ✅Filler Words
- ✅Words per min
- ✅Number of questions asked
- ✅Change in pace
- ✅Silence in Speech
- ✅Monotonous Pitch Detection
- ✅Volume
- ✅Emotional Detection
- ✅Eye Contact
- ❌Gesture Detection
- ❌CNN model for comparison

Note:-✅ completed, ❌ not in current scope, ䷢ in progress

### Version Release Information

#### v0.1

## How to use this software?

**Step 1**:
Download Repository and dependencies

```bash
conda create -n betterteaching python=3.7.3
```

```bash
conda activate betterteaching
```

```bash
git clone https://github.com/UnpredictablePrashant/BetterTeaching.git
```

```bash
cd BetterTeaching
```

```bash
python -m pip install -r requirements.txt
```

**Step 2**:
Modify the [credentials.json file](./config/credentials.json)

**Step 3**
Put your video `*.mp4` on `./videos/` folder

**Step 4**: Run python script in console to do the speech to text conversion
```bash
python app_transcribe.py
```

**Step 5**:
```bash
(BetterTeaching) x:\xxxx>python
Python 3.7.3 (default, Apr 24 2019, 15:29:51) [MSC v.1915 64 bit (AMD64)] :: Anaconda, Inc. on win32
Type "help", "copyright", "credits" or "license" for more information.
>>> from src import reset
>>> reset()
>>> exit()
```

**Step 6**: Run python script in console to do emotional analysis
```bash
python app_rekognition.py
```

**Step 7**: Run python script in console to generate the analysis report
```bash
python app_analysis.py
```

**Step 8**: Run python script in console to generate pdf analysis report
```bash
python generatePdf.py
```

Note:-
All the results will be available at `./out/` folder

## Author
Prashant Kumar Dey<br>
Ashish Kumar<br>
Ana Nikolic