# Image2Excel
Procrastination and silly ideas unite! A basic and slow image to excel converter (powered by ImageSharp)

Disclaimer: When I say it is slow, it's VERY slow. In the order of only 60-100 pixels per second without parallelization (which, luckily, yields improvements of about 80% in running times). As it turns out, OLE interop calls are sluggish, especially in the order of thousands or millions of cell iterations. Who would have guessed that?