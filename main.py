import requests
import json
import Constants
import win32com.client as wincom

speak = wincom.Dispatch("SAPI.SpVoice")

while True:
    city = input("Enter the name of the city (type 'exit' to quit): ")
    service_key = Constants.API_KEY_SERVICE

    if city.lower() == "exit":
        print("Weather app is exiting.")
        break

    url = f"http://api.weatherapi.com/v1/current.json?key={service_key}&q={city}"

    req = requests.get(url)

    if req.status_code == 200:
        weather_data = json.loads(req.text)

        if "error" not in weather_data:
            location = weather_data["location"]
            current = weather_data["current"]

            print(f"Weather in {location['name']}, {location['region']}, {location['country']}:")
            print(f"Temperature: {current['temp_c']}°C")
            print(f"Condition: {current['condition']['text']}")

            speak.Speak(f"The current weather in {city} is {current['temp_c']}°C {current['condition']['text']}")
        else:
            print("Error fetching weather data.")
            speak.Speak("Sorry, there is an error.")
    else:
        print("Error fetching weather data.")
        speak.Speak("Sorry, there is an error.")
