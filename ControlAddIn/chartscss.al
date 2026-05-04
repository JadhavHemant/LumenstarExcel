controladdin chartcss
{
    MinimumHeight = 200;
    MinimumWidth = 200;
    MaximumHeight = 300;
    RequestedHeight = 300;
    VerticalStretch = true;
    HorizontalStretch = true;
    StyleSheets = 'https://cdn.jsdelivr.net/npm/charts.css/dist/charts.min.css',
                  'chartscss.css';
    StartupScript = 'chartscss-startup.js';
    Scripts = 'chartscss-functions.js';

    procedure Render(html: Text);

    event ControlReady();
}