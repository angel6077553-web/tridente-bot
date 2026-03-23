const restify = require('restify');
const { BotFrameworkAdapter } = require('botbuilder');

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId || '',
    appPassword: process.env.MicrosoftAppPassword || ''
});

adapter.onTurnError = async (context, error) => {
    console.error('[onTurnError]', error);
    await context.sendActivity('Ocurrió un error. Intenta de nuevo.');
};

// Respuestas temáticas del bot
function getResponse(text) {
    const msg = text.toLowerCase().trim();

    if (/hola|hi|hey|buenas/.test(msg))
        return '👋 ¡Bienvenido al Tridente! Puedo contarte sobre el LaFerrari, el Porsche 918 Spyder y el McLaren P1. ¿Cuál te interesa?';

    if (/laferrari|ferrari/.test(msg))
        return '🔴 El LaFerrari es la obra maestra de Maranello. V12 6.3L + sistema HY-KERS de F1 = 963 CV. Solo 499 unidades producidas. 0-100 en 2.9s y más de 350 km/h. ¿Quieres saber más?';

    if (/918|porsche|spyder/.test(msg))
        return '🏆 El Porsche 918 Spyder rompió el récord de Nürburgring con 6:57. V8 4.6L + 2 motores eléctricos = 887 CV y 1.280 Nm de par. Solo 918 unidades. El más tecnológico del tridente.';

    if (/p1|mclaren/.test(msg))
        return '🟠 El McLaren P1 es F1 para la calle. V8 3.8L Twin-Turbo + motor eléctrico = 916 CV. 600 kg de downforce con aerodinámica activa. Solo 375 unidades. El más extremo del tridente.';

    if (/rapido|rápido|velocidad|speed|top speed/.test(msg))
        return '💨 Los tres superan 340 km/h. El LaFerrari y el P1 llegan a 350+ km/h. El 918 Spyder alcanza 345 km/h. Pero el más rápido de 0-100 es el 918 con 2.6 segundos gracias a su tracción total.';

    if (/potencia|hp|cv|caballos|horsepower/.test(msg))
        return '⚡ LaFerrari: 963 CV | McLaren P1: 916 CV | Porsche 918: 887 CV. Los tres son híbridos — combinan motor de combustión con eléctrico. El tridente suma 2.766 CV en total 🤯';

    if (/precio|costo|cuanto|vale|cost/.test(msg))
        return '💸 LaFerrari: ~$1.4M USD | McLaren P1: ~$1.15M USD | Porsche 918 Spyder: ~$845K USD. Todos agotados en semanas. Hoy en subasta pueden triplicar su valor.';

    if (/unidades|cuantos|produccion|producción/.test(msg))
        return '🏭 Producción limitadísima: LaFerrari = 499 ud. | McLaren P1 = 375 ud. | Porsche 918 = 918 ud. Total: solo 1.792 autos en el mundo.';

    if (/mejor|favorito|ganar|vs|versus|comparar/.test(msg))
        return '🤔 ¡La pregunta del millón! El 918 es el más versátil y el más rápido en Nürburgring. El P1 es el más extremo y enfocado. El LaFerrari es el más emotivo y puro. ¿Cuál es tu favorito?';

    if (/gracias|thanks/.test(msg))
        return '🏎️ ¡De nada! El Tridente siempre a tu servicio. ¿Algo más sobre estos hypercars legendarios?';

    if (/adios|bye|chau|hasta/.test(msg))
        return '👋 ¡Hasta pronto! Recuerda: LaFerrari · 918 Spyder · McLaren P1 — los tres hypercars que cambiaron la historia. 🏁';

    return '🏎️ Puedo contarte sobre el LaFerrari, el Porsche 918 Spyder o el McLaren P1. Pregúntame sobre velocidad, potencia, precio, producción o cuál es el mejor. ¿Qué quieres saber?';
}

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        if (context.activity.type === 'message') {
            const userText = context.activity.text || '';
            const reply = getResponse(userText);
            await context.sendActivity(reply);
        } else if (context.activity.type === 'conversationUpdate') {
            for (const member of context.activity.membersAdded || []) {
                if (member.id !== context.activity.recipient.id) {
                    await context.sendActivity(
                        '🏎️ ¡Hola! Puedo contarte todo sobre el LaFerrari, el Porsche 918 Spyder y el McLaren P1. ¿Por dónde empezamos?'
                    );
                }
            }
        }
    });
});

const PORT = process.env.PORT || 10000;
server.listen(PORT, () => {
    console.log(`Bot escuchando en puerto ${PORT}`);
});
