from pptx import Presentation
from pptx_tools import utils
import os
import edge_tts
import asyncio
import cv2
import math
import shutil
import pandas as pd
from mutagen.mp3 import MP3
from moviepy.editor import VideoFileClip, AudioFileClip, concatenate_videoclips

def text_to_mp3(text, mp3_filename):
    print(mp3_filename)
    voice = 'zh-CN-XiaoxiaoNeural'

    async def tts(text, mp3_filename):
        tts = edge_tts.Communicate(text=text, voice=voice)
        await tts.save(mp3_filename)

    asyncio.run(tts(text, mp3_filename))

def image_to_video(img_filename, video_filename, duration, fps):
    print(video_filename, duration)
    os.makedirs(os.path.dirname(video_filename), exist_ok=True)
    img_size = (1920, 1080)
    fourcc = cv2.VideoWriter_fourcc(*"mp4v")
    frame = cv2.imread(img_filename)
    frame = cv2.resize(frame, img_size)
    videoWriter = cv2.VideoWriter(video_filename, fourcc, fps, img_size)
    for i in range(0, duration):
        videoWriter.write(frame)
    videoWriter.release()

def get_mp3_duration(mp3_filename):
    audio = MP3(mp3_filename)
    time_count = audio.info.length
    return time_count

def add_sound_to_video(video_filename, mp3_filename, output_filename):
    print(output_filename)
    os.makedirs(os.path.dirname(output_filename), exist_ok=True)
    video = VideoFileClip(video_filename)
    videos = video.set_audio(AudioFileClip(mp3_filename)) 
    videos.write_videofile(output_filename, audio_codec='aac') 

def merge_vidio(video_files, output_filename):
    print(output_filename)
    os.makedirs(os.path.dirname(output_filename), exist_ok=True)
    clip_lst = []
    for video in video_files:
        clip_lst.append(VideoFileClip(video))
    final_clip = concatenate_videoclips(clip_lst)
    final_clip.write_videofile(output_filename, codec="libx264", fps=24, audio_bitrate="50k") 

def cvs_to_mp3(cvs_filename, mp3_pathname):
    os.makedirs(mp3_pathname, exist_ok=True)
    out_data = pd.read_csv(cvs_filename)
    text_all = ''
    for i in range(len(out_data)):
        text  = out_data['text'][i]
        text_all += text + '\n'
        mp3_filename = os.path.join(mp3_pathname, '{}.mp3'.format(i + 1))
        text_to_mp3(text, mp3_filename)

    mp3_filename = os.path.join(mp3_pathname, 'all.mp3')
    text_to_mp3(text_all, mp3_filename)
    return len(out_data)

def slide_to_image(ppt_filename, img_pathname):
    print(ppt_filename, img_pathname)
    if os.path.isdir(img_pathname):
        shutil.rmtree(img_pathname) 
    os.makedirs(img_pathname, exist_ok=True)

    ppt_filename = os.path.abspath(ppt_filename)
    img_pathname = os.path.abspath(img_pathname)
    utils.save_pptx_as_png(img_pathname, ppt_filename, overwrite_folder=True)

def make_video(res_pathname, temp_pathname, output_filename, fps = 24):

    cvs_filename = os.path.join(res_pathname, 'text.csv')
    mp3_pathname = os.path.join(temp_pathname, 'mp3')
    mp3_total = cvs_to_mp3(cvs_filename, mp3_pathname)

    img_pathname = os.path.join(temp_pathname, 'image')
    ppt_filename = os.path.join(res_pathname, 'slide.pptx')
    slide_to_image(ppt_filename, img_pathname)

    video_pathname = os.path.join(temp_pathname, 'video')
    video_list = []

    for i in range(1, mp3_total + 1):
        mp3_filename = os.path.join(mp3_pathname, '{}.mp3'.format(i))
        mp3_duration = get_mp3_duration(mp3_filename)
        mp3_duration = round(mp3_duration * fps)
        video_filename = os.path.join(video_pathname, '{}.mp4'.format(i))
        img_zh_filename = os.path.join(img_pathname, '幻灯片{}.PNG'.format(i))
        img_filename = os.path.join(img_pathname, '{}.PNG'.format(i))
        os.rename(img_zh_filename, img_filename)
        img_filename = os.path.abspath(img_filename)
        image_to_video(img_filename, video_filename, mp3_duration, fps)
        video_list.append(video_filename)
        
    mp4_all_filename = os.path.join(video_pathname, 'all.mp4')
    merge_vidio(video_list, mp4_all_filename)
    mp3_all_filename = os.path.join(mp3_pathname, 'all.mp3')
    add_sound_to_video(mp4_all_filename, mp3_all_filename, output_filename)
        

if __name__ == '__main__':
    
    make_video('./data/ciyi/aiqiu_kenqiu', './temp', './output/aiqiu_kenqiu.mp4')
