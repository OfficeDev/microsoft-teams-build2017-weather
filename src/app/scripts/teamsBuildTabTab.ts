import { TeamsTheme } from './theme';

// Changes for WeatherTab
interface mapLocation {
    zip: string,
    cityName: string,
    address: {
        adminDistrict: string,
        adminDistrict2: string,
        countryRegion: string,
        formattedAddress: string,
        locality: string
    },
    lat: string,
    lon:string
}
// Changes for WeatherTab

/**
 * Implementation of the Teams Build Tab content page
 */
export class teamsBuildTabTab {
    /**
     * Constructor for teamsBuildTab that initializes the Microsoft Teams script and themes management
     */
    constructor() {
        microsoftTeams.initialize();
    }
    
    /**
     * Method to invoke on page to start processing
     * Add your custom implementation here
     */
    public doStuff() {
        microsoftTeams.getContext((context: microsoftTeams.Context) => {
            /* Default code
            let a = document.getElementById('app');
            if (a) {
                a.innerHTML = `The value is: ${this.getParameterByName('data')}`;
            }
            */
            // Changes for WeatherTab
            const httpsCorsWrapper = "https://cors-anywhere.herokuapp.com/";
            const ipUrl = httpsCorsWrapper + "http://ip-api.com/json";
            const openWeatherAPIKey = "3979865d03301e5cf1887b80c64c2f9f";
            var openWeatherUrl = httpsCorsWrapper + `http://api.openweathermap.org/data/2.5/weather?appid=${openWeatherAPIKey}&units=imperial&`;
            const bingMapsAPIKey = "AhxX0G46a9kuxNIXI6YxGgh8w7aZdx4mAHkkVbYS5kQ_eNmxMKa8qYtBGvYv3Ohf";
            var bingMapsUrl: string;
            var weatherLoc = <mapLocation>{};
            $(document).ready(() => {
                let tabData = context.entityId;
                if (tabData) {
                    if (isNaN(Number(tabData))) {
                        // Assume it's a city name
                        openWeatherUrl = openWeatherUrl + `q=${tabData}`;
                    }
                    else {
                        openWeatherUrl = openWeatherUrl + `zip=${tabData},us`;
                        weatherLoc.zip = tabData.trim();
                    }
                }
                else {
                    // Get the current location based on IP address
                    $.getJSON(ipUrl, (locate: any) => {
                        openWeatherUrl = openWeatherUrl + `lat=${locate.lat}&lon=${locate.lon}`;
                        weatherLoc.lat = locate.lat;
                        weatherLoc.lon = locate.lon;
                    });
                }
                // Retrieve current weather conditions for the city, zip code, or lat/lon
                $.getJSON(openWeatherUrl, (data: any) => {
                    var code = data.weather[0].id;
                    var hour = (new Date()).getHours();
                    var darkOutside = function nightOrDay(conditionCode:string) {
                        if ((hour > 18) || (hour < 7))  {
                            return conditionCode + '-n';
                        }
                        else {
                            return conditionCode + '-d';
                        }
                    }
                    // Set lat/long if not already known
                    if (!weatherLoc.lat) {
                        weatherLoc.lat = data.coord.lat;
                        weatherLoc.lon = data.coord.lon;
                    }
                    // Retrieve location from Bing Maps reverse geocoding and update the UI
                    bingMapsUrl = `https://dev.virtualearth.net/REST/v1/Locations/${weatherLoc.lat},${weatherLoc.lon}?key=${bingMapsAPIKey}`;
                    $.ajax({
                        url: bingMapsUrl,
                        dataType: "jsonp",
                        jsonp: "jsonp",
                        success: ((locData: any) => {
                            weatherLoc.address = locData.resourceSets[0].resources[0].address;
                            var weatherIconElem = <HTMLElement>document.getElementById("weathericon");
                            var weatherElem = <HTMLElement>document.getElementById("weather");
                            var tempElem = <HTMLElement>document.getElementById("temp");
                            var locElem = <HTMLElement>document.getElementById("location");
                            if (weatherIconElem && weatherElem && tempElem && locElem) {
                                weatherIconElem.className = "owf owf-" + darkOutside(code);
                                weatherElem.innerHTML = data.weather[0].main;
                                tempElem.innerHTML = String(Number(data.main.temp).toFixed(1)) + "&#176;";
                                locElem.innerHTML = `${weatherLoc.address.locality}, ${weatherLoc.address.adminDistrict}`;
                            }
                        }),
                        error: ((e: any) => {
                            console.log("Error calling Bing Maps reverse geocoding API");
                        })
                    })
                })
                // End changes for WeatherTab
            })
        })
    }

    /**
     * Method for retrieving query string parameters
     */
    private getParameterByName(name: string, url?: string): string {
        if (!url) {
            url = window.location.href;
        }
        name = name.replace(/[\[\]]/g, "\\$&");
        var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
            results = regex.exec(url);
        if (!results) return '';
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, " "));
    }

}