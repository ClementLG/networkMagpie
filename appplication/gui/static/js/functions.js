function start_button() {
    console.log("start_button")
    const location = window.location.hostname;
    const settings = {
        method: 'POST',
    };
    try {
        const fetchResponse = fetch(`http://${location}:5555/api/start`, settings);
    } catch (e) {
        return e;
    }
}

async function stop_button() {
    console.log("start_button")
    const location = window.location.hostname;
    const settings = {
        method: 'POST',
    };
    try {
        const fetchResponse = await fetch(`http://${location}:5555/api/stop`, settings);
    } catch (e) {
        return e;
    }
}