# deckflicks
Turns your Powerpoint deck into a narrated video using Azure text-to-speech

# requirements
pip install Spire.Presentation
pip install requests
pip install azure-cognitiveservices-speech

# command line
python deckflicks.py -s < azure speech subscription > -i < input file > -o < output file > -v < voice name >
