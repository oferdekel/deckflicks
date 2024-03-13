from spire.presentation.common import *
from spire.presentation import *
import azure.cognitiveservices.speech as speechsdk
import tempfile
import shutil
import requests
import argparse


def get_speech_voice_list(subscription_key):
    url = 'https://eastus.tts.speech.microsoft.com/cognitiveservices/voices/list'
    headers = {
        'Ocp-Apim-Subscription-Key': subscription_key
    }
    response = requests.get(url, headers=headers)
    text = str(response.text)
    return text


def text_to_wav(subscription_key, text, voice_name, wav_path):

    speech_config = speechsdk.SpeechConfig(subscription=subscription_key, region="eastus")
    speech_config.speech_synthesis_voice_name = voice_name
    audio_format = speechsdk.audio.AudioStreamFormat(samples_per_second=22050, bits_per_sample=16, channels=1)
    audio_config = speechsdk.AudioConfig(filename=wav_path)
    audio_config.stream_format=audio_format # TODO this doesn't do anything
    speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)
    
    result = speech_synthesizer.speak_text(text)

    if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
        return True
    elif result.reason == speechsdk.ResultReason.Canceled:
        cancellation_details = result.cancellation_details
        print("Speech synthesis canceled: {}".format(cancellation_details.reason))
        if cancellation_details.reason == speechsdk.CancellationReason.Error:
            print("Error details: {}".format(cancellation_details.error_details))
        return False


def add_speech_to_ppt(subscription_key, ppt_input_filename, ppt_output_filename, voice_name):
    presentation = Presentation()
    print(f'reading ppt file {ppt_input_filename}')
    presentation.LoadFromFile(ppt_input_filename)
    audios = presentation.WavAudios
    
    tempdir = tempfile.mkdtemp()
    print(f"creating temporary files in {tempdir}")

    for i, slide in enumerate(presentation.Slides):
        notesSlide = slide.NotesSlide
        text = notesSlide.NotesTextFrame.Text
        wav_path = f"{tempdir}/out{i}.wav"
        print(f'slide {i}, audio file: {wav_path}, text: {text[:80]}')

        text_to_wav(subscription_key, text, voice_name, wav_path)
        stream = Stream(wav_path)
        audioData = audios.Append(stream)
        rect = slide.Shapes.AppendAudioMedia("", RectangleF.FromLTRB(0, 0, 1, 1))
        rect.Data = audioData
        rect.Volume = AudioVolumeType.Loud

    print(f"removing {tempdir}")
    shutil.rmtree(tempdir)
    print(f'writing ppt file {ppt_output_filename}')
    presentation.SaveToFile(ppt_output_filename, FileFormat.Pptx2019)
    presentation.Dispose()


def main():
    parser = argparse.ArgumentParser(description='Turn your PPT into a video')
    parser.add_argument('-s', '--subscription_key', type=str, help='Azure speech subscription key', required=True)
    parser.add_argument('-i', '--ppt_input_filename', type=str, help='Input Powerpoint file', default='test.pptx')
    parser.add_argument('-o', '--ppt_output_filename', type=str, help='Output Powerpoint file', default='out.pptx')
    parser.add_argument('-v', '--voice_name', type=str, help='Output Powerpoint file', default='en-GB-RyanNeural')
    args = parser.parse_args()
    
    add_speech_to_ppt(args.subscription_key, args.ppt_input_filename, args.ppt_output_filename, args.voice_name)    

if __name__ == '__main__':
    main()

