import { useEffect, useState } from 'react'
import './App.css'

interface WeatherForecast {
  date: string
  temperatureC: number
  temperatureF: number
  summary: string | null
}

function App() {
  const [forecasts, setForecasts] = useState<WeatherForecast[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  function load() {
    setLoading(true)
    setError(null)
    fetch('/weatherforecast')
      .then((res) => {
        if (!res.ok) throw new Error(`API responded with ${res.status}`)
        return res.json() as Promise<WeatherForecast[]>
      })
      .then((data) => setForecasts(data))
      .catch((err: unknown) =>
        setError(err instanceof Error ? err.message : 'Failed to load weather'),
      )
      .finally(() => setLoading(false))
  }

  useEffect(load, [])

  const current = forecasts[0]

  return (
    <>
      <section id="center">
        <div>
          <h1>Current Weather</h1>

          {loading && <p>Loading weather…</p>}

          {error && (
            <p className="error">
              Could not reach the weather API: {error}. Make sure the WordApi
              backend is running on http://localhost:5269.
            </p>
          )}

          {current && (
            <div className="weather-now">
              <div className="temp">{current.temperatureC}°C</div>
              <div className="details">
                <p className="summary">{current.summary}</p>
                <p>{current.temperatureF}°F</p>
                <p className="date">
                  {new Date(current.date).toLocaleDateString(undefined, {
                    weekday: 'long',
                    month: 'long',
                    day: 'numeric',
                  })}
                </p>
              </div>
            </div>
          )}
        </div>

        <button type="button" className="counter" onClick={load}>
          Refresh
        </button>
      </section>

      <div className="ticks"></div>

      {forecasts.length > 1 && (
        <section id="next-steps">
          <div id="docs">
            <h2>Forecast</h2>
            <ul className="forecast">
              {forecasts.slice(1).map((f) => (
                <li key={f.date}>
                  <span className="date">
                    {new Date(f.date).toLocaleDateString(undefined, {
                      weekday: 'short',
                    })}
                  </span>
                  <span className="summary">{f.summary}</span>
                  <span className="temp">{f.temperatureC}°C</span>
                </li>
              ))}
            </ul>
          </div>
        </section>
      )}

      <div className="ticks"></div>
      <section id="spacer"></section>
    </>
  )
}

export default App
