from openai import OpenAI
import time

                
time.sleep(2)
completion = client.chat.completions.create(
    model="gpt-3.5-turbo-1106",
    messages=[
        {"role":"system","content":"you are a virtual assistent named Jarvis,skilled in general tasks like Alexa and Google."},
        {"role":"user","content":"what is coding."}
    ]
)
print(completion.choices[0].messages)
