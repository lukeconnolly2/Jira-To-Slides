1. Clone this repository
2. Install the dependencies
```bash
pip install -r requirements.txt
```
 
## Usage
1. Get your teams rapidViewId from the url of your Jira board. Example: /secure/RapidBoard.jspa?rapidView=**123**
2. Get your Jira JSESSIONID from your browser cookies.
3. Fill in the config.ini file.
 
```bash
python3 generate_slides.py
```
 
4. In the sprint demo slide deck, click on the beginning of your teams slides.
5. Click file > Import slides > Upload from file > Select the generated slides file > **Uncheck Keep Original Theme**
6. Click Import Slides
 
