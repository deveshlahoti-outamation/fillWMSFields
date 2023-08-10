import os
import sys
import openai
import time

openai.api_key = "cdd1176dbd164a49b4111e66f897e8fc"
openai.api_base = "https://wmsfieldfilling.openai.azure.com/"
openai.api_type = 'azure'
openai.api_version = "2023-03-15-preview"
deployment_name = 'WMSFieldFill'

# Get the directory of the .exe file
exe_dir = os.path.dirname(sys.executable)

# Build the path to your text file
user_prompt = os.path.join(exe_dir, 'user_prompt.txt')
system_prompt = os.path.join(exe_dir, 'system_prompt.txt')


def time_constrained_function(func, *args, timeout=10):
    start = time.time()
    result = func(*args)

    # If the function call exceeds the timeout
    if time.time() - start > timeout:
        print(f"Function execution exceeded {timeout} seconds, calling again...")
        return time_constrained_function(func, *args, timeout=timeout)

    return result


def format_comments(comments):
    for num in range(len(comments)):
        comments[num] = "Comment " + str(num + 1) + ": " + comments[num]
    return comments


def create_prompt(comments):
    prompt = (open(user_prompt)).read() + "\nActual Input:"
    for comment in format_comments(comments):
        prompt += "\n" + comment
    return prompt


def get_data(comments):
    try:
        prompt = create_prompt(comments)
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=[
                {"role": "system", "content": (open(system_prompt)).read()},
                {"role": "user", "content": prompt},
            ]
        )
        output = response['choices'][0]['message']['content']
        print(output)
        return output.split("; ")

    except Exception as e:
        print(e)
        get_data(comments)
