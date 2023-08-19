$APIKey = $env:OPENWEATHER_APIKEY

$Script:directions = @(
    @{ Name = "N"; Arrow = [char]::ConvertFromUtf32(0x2193) },
    @{ Name = "NNE"; Arrow = [char]::ConvertFromUtf32(0x2199) },
    @{ Name = "NE"; Arrow = [char]::ConvertFromUtf32(0x2199) },
    @{ Name = "ENE"; Arrow = [char]::ConvertFromUtf32(0x2199) },
    @{ Name = "E"; Arrow = [char]::ConvertFromUtf32(0x2190) },
    @{ Name = "ESE"; Arrow = [char]::ConvertFromUtf32(0x2196) },
    @{ Name = "SE"; Arrow = [char]::ConvertFromUtf32(0x2196) },
    @{ Name = "SSE"; Arrow = [char]::ConvertFromUtf32(0x2196) },
    @{ Name = "S"; Arrow = [char]::ConvertFromUtf32(0x2191) },
    @{ Name = "SSW"; Arrow = [char]::ConvertFromUtf32(0x2197) },
    @{ Name = "SW"; Arrow = [char]::ConvertFromUtf32(0x2197) },
    @{ Name = "WSW"; Arrow = [char]::ConvertFromUtf32(0x2197) },
    @{ Name = "W"; Arrow = [char]::ConvertFromUtf32(0x2192) },
    @{ Name = "WNW"; Arrow = [char]::ConvertFromUtf32(0x2198) },
    @{ Name = "NW"; Arrow = [char]::ConvertFromUtf32(0x2198) },
    @{ Name = "NNW"; Arrow = [char]::ConvertFromUtf32(0x2198) },
    @{ Name = "N"; Arrow = [char]::ConvertFromUtf32(0x2193) }
)
Function Convert-WindDirection {
    Param (
        [double]$Degree
    )
    $index = [Math]::Round(($Degree % 360) / 22.5)
    return $Script:directions[$index]
}
$Script:MoonPhases = @(
    @{ Name = "New Moon"; Icon = [char]::ConvertFromUtf32(0x1F311) },
    @{ Name = "Waxing Crescent"; Icon = [char]::ConvertFromUtf32(0x1F312) },
    @{ Name = "First Quarter"; Icon = [char]::ConvertFromUtf32(0x1F313) },
    @{ Name = "Waxing Gibbous"; Icon = [char]::ConvertFromUtf32(0x1F314) },
    @{ Name = "Full Moon"; Icon = [char]::ConvertFromUtf32(0x1F315) },
    @{ Name = "Waning Gibbous"; Icon = [char]::ConvertFromUtf32(0x1F316) },
    @{ Name = "Last Quarter"; Icon = [char]::ConvertFromUtf32(0x1F317) },
    @{ Name = "Waning Crescent"; Icon = [char]::ConvertFromUtf32(0x1F318) }
)
Function Convert-MoonPhase {
    Param (
        [double]$PhaseValue
    )

    if ($PhaseValue -ge 1) { $PhaseValue = 0 }

    $index = [Math]::Floor($PhaseValue * 8)
    return $Script:MoonPhases[$index]
}

$Emoji = @{
    Good     = [char]::ConvertFromUtf32(0x1F7E2)
    Fair     = [char]::ConvertFromUtf32(0x1F7E1)
    Moderate = [char]::ConvertFromUtf32(0x1F7E0)
    Poor     = [char]::ConvertFromUtf32(0x1F534)
    VeryPoor = [char]::ConvertFromUtf32(0x2620)
}
Function Convert-AirQualityIndex {
    Param ([int]$Value)
    switch ($Value) {
        1 { return "$($Emoji.Good) Good (1)" }
        2 { return "$($Emoji.Fair) Fair (2)" }
        3 { return "$($Emoji.Moderate) Moderate (3)" }
        4 { return "$($Emoji.Poor) Poor (4)" }
        5 { return "$($Emoji.VeryPoor)  Very Poor (5)" }
        default { return "Invalid input" }
    }
}
Function Convert-CloudCoverageToEmoji {
    Param ([int]$Coverage)

    switch ($Coverage) {
        { $_ -in 0..20 } { return [char]::ConvertFromUtf32(0x2600) }      # ☀️
        { $_ -in 21..40 } { return [char]::ConvertFromUtf32(0x1F324) }    # 🌤️
        { $_ -in 41..60 } { return [char]::ConvertFromUtf32(0x26C5) }     # ⛅
        { $_ -in 61..80 } { return [char]::ConvertFromUtf32(0x1F325) }    # 🌥️
        { $_ -in 81..100 } { return [char]::ConvertFromUtf32(0x2601) }    # ☁️
        default { return "Invalid cloud coverage value" }
    }
}

function Convert-AirQuality($type, $value) {
    $thresholds = @{
        "PM25" = @{"good" = 0; "fair" = 10; "moderate" = 25; "poor" = 50; "veryPoor" = 75 };
        "PM10" = @{"good" = 0; "fair" = 20; "moderate" = 50; "poor" = 100; "veryPoor" = 200 };
        "SO2"  = @{"good" = 0; "fair" = 20; "moderate" = 80; "poor" = 250; "veryPoor" = 350 };
        "NO2"  = @{"good" = 0; "fair" = 40; "moderate" = 70; "poor" = 150; "veryPoor" = 200 };
        "O3"   = @{"good" = 0; "fair" = 60; "moderate" = 100; "poor" = 140; "veryPoor" = 180 };
        "CO"   = @{"good" = 0; "fair" = 4400; "moderate" = 9400; "poor" = 12400; "veryPoor" = 15400 };
        "NH3"  = @{"good" = 0.1; "fair" = 50; "moderate" = 100; "poor" = 150; "veryPoor" = 200 };
        "NO"   = @{"good" = 0.1; "fair" = 25; "moderate" = 50; "poor" = 75; "veryPoor" = 100 };
    }

    $currentThreshold = $thresholds[$type]
    $quality = ""

    switch ($value) {
        { $_ -ge $currentThreshold["good"] -and $_ -lt $currentThreshold["fair"] } { 
            $quality = "$($Emoji["Good"]) $value - Good ($($currentThreshold["good"])-$($currentThreshold["fair"]))" 
        }
        { $_ -ge $currentThreshold["fair"] -and $_ -lt $currentThreshold["moderate"] } { 
            $quality = "$($Emoji["Fair"]) $value - Fair ($($currentThreshold["fair"])-$($currentThreshold["moderate"]))"
        }
        { $_ -ge $currentThreshold["moderate"] -and $_ -lt $currentThreshold["poor"] } { 
            $quality = "$($Emoji["Moderate"]) $value - Moderate ($($currentThreshold["moderate"])-$($currentThreshold["poor"]))" 
        }
        { $_ -ge $currentThreshold["poor"] -and $_ -lt $currentThreshold["veryPoor"] } { 
            $quality = "$($Emoji["Poor"]) $value - Poor ($($currentThreshold["poor"])-$($currentThreshold["veryPoor"]))"
        }
        { $_ -ge $currentThreshold["veryPoor"] } { 
            $quality = "$($Emoji["VeryPoor"])  $value - Very Poor ($($currentThreshold["veryPoor"])+)" 
        }
    }    
    return $quality
}
$UVEmoji = @{
    Green  = [char]::ConvertFromUtf32(0x1F7E2)
    Yellow = [char]::ConvertFromUtf32(0x1F7E1)
    Orange = [char]::ConvertFromUtf32(0x1F7E0)
    Red    = [char]::ConvertFromUtf32(0x1F534)
    Purple = [char]::ConvertFromUtf32(0x1F7E3)
}

Function Get-UVIndexMessage {
    Param ([double]$UVIndex)

    $RoundedUVIndex = [math]::Round($UVIndex)

    switch ($RoundedUVIndex) {
        { $_ -in 0..2 } { return "$($UVEmoji.Green) $UVIndex - Low" }
        { $_ -in 3..5 } { return "$($UVEmoji.Yellow) $UVIndex - Moderate" }
        { $_ -in 6..7 } { return "$($UVEmoji.Orange) $UVIndex - High" }
        { $_ -in 8..10 } { return "$($UVEmoji.Red) $UVIndex - Very High" }
        { $_ -ge 11 } { return "$($UVEmoji.Purple) $UVIndex - Extreme" }
        default { return "Invalid UV index value" }
    }
}




Function Invoke-API {
    Param (
        [string]$Uri
    )
    $Response = Invoke-RestMethod -Uri $Uri

    return $Response
}
Function Get-WeatherData {
    Param (
        [string]$City = '',
        [ValidateSet('Imperial', 'Metric')]
        [string]$Units = 'Imperial'
    )

    if ($City) {
        $Response_Weather = Invoke-API -Uri "http://api.openweathermap.org/data/2.5/weather?q=$City&appid=$APIKey&units=$Units"
    }
    else {
        $Response_Weather = Invoke-API -Uri "http://api.openweathermap.org/data/2.5/weather?lat=40.4661601&lon=-80.0319426&appid=$APIKey&units=$Units"
    }

    $Lat = $Response_Weather.coord.lat
    $Lon = $Response_Weather.coord.lon
    
    $Response_OneCall = Invoke-API -Uri "https://api.openweathermap.org/data/2.5/onecall?lat=$Lat&lon=$Lon&exclude=minutely,alerts&appid=$APIKey&units=$Units"
    $Response_AirPollution = Invoke-API -Uri "http://api.openweathermap.org/data/2.5/air_pollution?lat=$Lat&lon=$Lon&appid=$APIKey"
    
    $WindDirection = Convert-WindDirection -Degree $Response_OneCall.current.wind_deg
    $WindGust = if ($Response_OneCall.current.wind_gust) {
        $Response_OneCall.current.wind_gust.ToString() + " " + $SpeedUnit
    }
    else {
        "N/A"
    }

    $CloudCoverageEmoji = Convert-CloudCoverageToEmoji -Coverage $Response_OneCall.current.clouds

    $MoonPhase = Convert-MoonPhase -PhaseValue $Response_OneCall.daily[0].moon_phase
    
    $TempUnit = if ($Units -eq 'Imperial') { "F" } else { "C" }
    $SpeedUnit = if ($Units -eq 'Imperial') { "mph" } else { "m/s" }

    $UVIndexMessage = Get-UVIndexMessage -UVIndex $Response_OneCall.current.uvi

    $AirQualityIndex = Convert-AirQualityIndex -Value $Response_AirPollution.list.main.aqi

    $PM25Quality = Convert-AirQuality "PM25" $Response_AirPollution.list[0].components.pm2_5
    $PM10Quality = Convert-AirQuality "PM10" $Response_AirPollution.list[0].components.pm10
    $SO2Quality = Convert-AirQuality "SO2" $Response_AirPollution.list[0].components.so2
    $NO2Quality = Convert-AirQuality "NO2" $Response_AirPollution.list[0].components.no2
    $O3Quality = Convert-AirQuality "O3" $Response_AirPollution.list[0].components.o3
    $COQuality = Convert-AirQuality "CO" $Response_AirPollution.list[0].components.co
    $NH3AirQuality = Convert-AirQuality "NH3" $Response_AirPollution.list[0].components.nh3
    $NOAirQuality = Convert-AirQuality "NO" $Response_AirPollution.list[0].components.no

    $WeatherData = [PSCustomObject]@{
        "Location Name"            = $Response_Weather.name
        "Weather Description"      = $Response_OneCall.current.weather.description
        "Current Temperature"      = $Response_OneCall.current.temp.ToString() + " " + $degreeSymbol + $TempUnit
        "Feels Like Temperature"   = $Response_OneCall.current.feels_like.ToString() + " " + $degreeSymbol + $TempUnit
        "Minimum Temperature"      = $Response_OneCall.daily[0].temp.min.ToString() + " " + $degreeSymbol + $TempUnit
        "Maximum Temperature"      = $Response_OneCall.daily[0].temp.max.ToString() + " " + $degreeSymbol + $TempUnit
        "UV Index"                 = $UVIndexMessage
        "Cloud Coverage"           = "$CloudCoverageEmoji  $($Response_OneCall.current.clouds)%"        
        "Precipitation"            = ($Response_OneCall.hourly[0].pop * 100).ToString() + '%'
        "Dew Point"                = $Response_OneCall.current.dew_point.ToString() + " " + $degreeSymbol + $TempUnit
        "Humidity"                 = $Response_OneCall.current.humidity.ToString() + '%'
        "Pressure"                 = $Response_OneCall.current.pressure.ToString() + " hPa"
        "Visibility"               = $Response_OneCall.current.visibility.ToString() + " km"
        "Wind Speed"               = $Response_OneCall.current.wind_speed.ToString() + " " + $SpeedUnit
        "Wind Gust"                = $WindGust
        "Wind Direction"           = $WindDirection.Name + " " + $WindDirection.Arrow 
        "Air Quality Index"        = $AirQualityIndex
        "Air Pollution (PM2.5)"    = $PM25Quality
        "Air Pollution (PM10)"     = $PM10Quality
        "Air Pollution (SO2)"      = $SO2Quality
        "Air Pollution (NO2)"      = $NO2Quality
        "Air Pollution (O3)"       = $O3Quality
        "Air Pollution (CO)"       = $COQuality
        "Air Pollution (NH3)"      = $NH3AirQuality
        "Air Pollution (NO)"       = $NOAirQuality
        "Sunrise Time"             = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($Response_OneCall.daily[0].sunrise)).ToString("h:mm:ss tt M/d/yyyy")
        "Sunset Time"              = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($Response_OneCall.daily[0].sunset)).ToString("h:mm:ss tt M/d/yyyy")
        "Moon Phase"               = $MoonPhase.Name + " " + $MoonPhase.icon
        "Coordinates"              = "$Lat,$Lon"
        "Google Maps"              = "https://www.google.com/maps/?q=$Lat,$Lon"
        "Time of Data Calculation" = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($Response_OneCall.current.dt)).ToString("h:mm:ss tt M/d/yyyy")
    }    
    
    # Output the custom object
    return $WeatherData 
}
