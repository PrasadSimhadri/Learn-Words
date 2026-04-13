import * as XLSX from 'xlsx';

// Initial data extracted from the user's provided vocab.xlsx
const INITIAL_DATA = [
    {"Word":"betray","Meaning":"to give information about somebody/something to an enemy; to make a secret known.","Example":"She betrayed all the members of the group to the secret police."},
    {"Word":"ambigious","Meaning":"Having more than one possible meaning; unclear or open to different interpretations.","Example":"he ending of the movie was ambiguous, leaving the audience guessing."},
    {"Word":"obscure","Meaning":"Not well known; unclear or hard to understand","Example":"The scientist studied an obscure species of insect that few people had ever heard of."},
    {"Word":"undermine","Meaning":"to weaken or damage something gradually","Example":"The rumors began to undermine the company's reputation"},
    {"Word":"commensurate","Meaning":"to be equal in size, amount ","Example":"his salary is commensurate with this experience and skills"},
    {"Word":"venality","Meaning":"being open to bribe or corruption, willing to act dishonestly for money or personal gain","Example":"people lost trust in the system due with widespread venality"},
    {"Word":"egregious","Meaning":"outstandily bad, shocking or extremely wrong in a noticeable way","Example":"the student made an egregious mistake by using phone during exam"},
    {"Word":"upbraid","Meaning":"to scold or criticize hardly, specially for doing something wrong ","Example":"the teacher upbraided the student for cheating on the test"},
    {"Word":"calumny","Meaning":"false statements / lies made to damage one's reputation","Example":"politician denied the accusations, calling them pure calumny"},
    {"Word":"venerate","Meaning":"to respect or honor someone or something deeply","Example":"many people venerate great leaders for their contributions"},
    {"Word":"disinterested","Meaning":"means impartial, unbiased, uninvested in an outcome","Example":"A disinterested party was needed to mediate the dispute"},
    {"Word":"ameliorate","Meaning":"making something better/improve a bad situation","Example":"I ameliorated by friend's mood after a rough day"},
    {"Word":"mitigate","Meaning":"to reduce seriousness/bad effects","Example":"conversation can mitigate misunderstanding"},
    {"Word":"reticent","Meaning":"one who doesn’t talk much","Example":"he is very reticent in the class"},
    {"Word":"pragmatic","Meaning":"practical and realistic","Example":"he takes very pragmatic decisions about his future"},
    {"Word":"equivocal","Meaning":"saying something unclear/open to interpretation","Example":"his equivocal answers made him a prime suspect for the officers"},
    {"Word":"ephemeral","Meaning":"something that’s short lived/for short time/temporary","Example":"his happiness was ephemeral"},
    {"Word":"taciturn","Meaning":"person who's very quite and speaks very little","Example":"he remained taciturn during the meeting"},
    {"Word":"duplicity","Meaning":"someone dishonest and pretends to be truthful while hiding truth","Example":"his duplicity shocked everyone"}
];

let words = [...INITIAL_DATA];
let currentIndex = 0;
let isFlipped = false;

// DOM Elements
const card = document.getElementById('card');
const wordDisplay = document.getElementById('word-display');
const meaningDisplay = document.getElementById('meaning-display');
const exampleDisplay = document.getElementById('example-display');
const flipBtn = document.getElementById('flip-btn');
const nextBtn = document.getElementById('next-btn');
const prevBtn = document.getElementById('prev-btn');
const counter = document.getElementById('counter');
const progressFill = document.getElementById('progress-fill');
const shuffleBtn = document.getElementById('shuffle-btn');
const excelUpload = document.getElementById('excel-upload');
const resetBtn = document.getElementById('reset-btn');

function shuffle(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
}

function updateCard() {
    if (words.length === 0) {
        wordDisplay.textContent = "No Words Loaded";
        meaningDisplay.textContent = "";
        exampleDisplay.textContent = "";
        return;
    }

    const currentWord = words[currentIndex];
    
    // Add transition effect
    card.classList.add('fading');
    
    setTimeout(() => {
        wordDisplay.textContent = currentWord.Word || "N/A";
        meaningDisplay.textContent = currentWord.Meaning || "No meaning provided.";
        exampleDisplay.textContent = currentWord.Example || "";
        
        // Update stats
        counter.textContent = `${currentIndex + 1} / ${words.length}`;
        const progress = ((currentIndex + 1) / words.length) * 100;
        progressFill.style.width = `${progress}%`;
        
        card.classList.remove('fading');
    }, 200);

    // Reset flip state when moving to new word
    if (isFlipped) {
        toggleFlip();
    }
}

function toggleFlip() {
    isFlipped = !isFlipped;
    card.classList.toggle('is-flipped', isFlipped);
    flipBtn.textContent = isFlipped ? "Show Word" : "Show Meaning";
}

function nextWord() {
    if (words.length === 0) return;
    currentIndex = (currentIndex + 1) % words.length;
    updateCard();
}

function prevWord() {
    if (words.length === 0) return;
    currentIndex = (currentIndex - 1 + words.length) % words.length;
    updateCard();
}

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        
        if (json.length > 0) {
            words = shuffle(json);
            // Save to local storage for persistence
            localStorage.setItem('wordwise_data', JSON.stringify(json));
            
            currentIndex = 0;
            updateCard();
            alert(`Loaded ${json.length} words successfully and saved to local storage!`);
        } else {
            alert("No words found in the Excel file.");
        }
    };
    reader.readAsArrayBuffer(file);
}

// Event Listeners
flipBtn.addEventListener('click', toggleFlip);
card.addEventListener('click', toggleFlip);
nextBtn.addEventListener('click', nextWord);
prevBtn.addEventListener('click', prevWord);
shuffleBtn.addEventListener('click', () => {
    words = shuffle([...words]);
    currentIndex = 0;
    updateCard();
});
excelUpload.addEventListener('change', handleFileUpload);
resetBtn.addEventListener('click', () => {
    if (confirm("Are you sure you want to reset to default words and clear saved data?")) {
        localStorage.removeItem('wordwise_data');
        words = [...INITIAL_DATA];
        words = shuffle([...words]);
        currentIndex = 0;
        updateCard();
    }
});

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    if (e.code === 'Space') {
        e.preventDefault();
        toggleFlip();
    } else if (e.code === 'ArrowRight') {
        nextWord();
    } else if (e.code === 'ArrowLeft') {
        prevWord();
    }
});

// Initialize
function init() {
    const savedData = localStorage.getItem('wordwise_data');
    if (savedData) {
        try {
            words = JSON.parse(savedData);
            console.log('Loaded words from localStorage');
        } catch (e) {
            console.error('Failed to parse saved data', e);
            words = [...INITIAL_DATA];
        }
    } else {
        words = [...INITIAL_DATA];
    }
    words = shuffle([...words]);
    updateCard();
}

init();
