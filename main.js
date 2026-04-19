import * as XLSX from 'xlsx';

// Initial data for welcome state
const INITIAL_DATA = [
    {"Word":"Upload and learn","Meaning":"Please upload an Excel file to start learning new words.","Example":"Click the 'Load Excel' button to get started!", "Synonyms": "Progress, Growth, Mastery"}
];

let fullWordList = [...INITIAL_DATA]; // Holds all loaded words
let words = [...INITIAL_DATA]; // Holds the currently displayed set (can be a range)
let currentWorkbook = null; // Store workbook to switch sheets without re-uploading
let currentIndex = 0;
let isFlipped = false;

// DOM Elements
const card = document.getElementById('card');
const wordDisplay = document.getElementById('word-display');
const meaningDisplay = document.getElementById('meaning-display');
const exampleDisplay = document.getElementById('example-display');
const synonymsDisplay = document.getElementById('synonyms-display');
const synonymsGroup = document.getElementById('synonyms-group');
const flipBtn = document.getElementById('flip-btn');
const nextBtn = document.getElementById('next-btn');
const prevBtn = document.getElementById('prev-btn');
const counter = document.getElementById('counter');
const progressFill = document.getElementById('progress-fill');
const shuffleBtn = document.getElementById('shuffle-btn');
const excelUpload = document.getElementById('excel-upload');
const resetBtn = document.getElementById('reset-btn');

// Range Controls
const rangeFromInput = document.getElementById('range-from');
const rangeToInput = document.getElementById('range-to');
const loadRangeBtn = document.getElementById('load-range-btn');
const sheetSelectorGroup = document.getElementById('sheet-selector-group');
const sheetSelect = document.getElementById('sheet-select');
const bookmarkBtnFront = document.getElementById('bookmark-btn-front');
const bookmarkBtnBack = document.getElementById('bookmark-btn-back');

function shuffle(array) {
    const newArray = [...array];
    for (let i = newArray.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
    }
    return newArray;
}

function updateCard() {
    if (words.length === 0) {
        wordDisplay.textContent = "No Words Loaded";
        meaningDisplay.textContent = "";
        exampleDisplay.textContent = "";
        synonymsDisplay.textContent = "";
        counter.textContent = "0 / 0";
        progressFill.style.width = "0%";
        return;
    }

    const currentWord = words[currentIndex];
    
    // Add transition effect
    card.classList.add('fading');
    
    setTimeout(() => {
        wordDisplay.textContent = currentWord.Word || "N/A";
        const meaningText = currentWord.Meaning || "No meaning provided.";
        meaningDisplay.textContent = meaningText;
        exampleDisplay.textContent = currentWord.Example || "";
        
        const synonyms = currentWord.Synonyms || "";
        synonymsDisplay.textContent = synonyms;
        synonymsGroup.style.display = synonyms ? 'block' : 'none';
        
        // Update stats
        counter.textContent = `${currentIndex + 1} / ${words.length}`;
        const progress = ((currentIndex + 1) / words.length) * 100;
        progressFill.style.width = `${progress}%`;
        
        // Update bookmark buttons
        const isBookmarked = currentWord.isBookmarked || false;
        bookmarkBtnFront.classList.toggle('active', isBookmarked);
        bookmarkBtnBack.classList.toggle('active', isBookmarked);
        
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

function bookmarkWord(event) {
    if (event) event.stopPropagation(); // Prevent card flip
    if (words.length === 0) return;

    const currentWord = words[currentIndex];
    
    // Find the word in fullWordList to update its state permanently
    const wordInFullList = fullWordList.find(w => w.Word === currentWord.Word);
    
    // Toggle state
    const newBookmarkState = !currentWord.isBookmarked;
    currentWord.isBookmarked = newBookmarkState;
    if (wordInFullList) wordInFullList.isBookmarked = newBookmarkState;

    if (newBookmarkState) {
        // Shown again after at random point of time along of other words left
        // "Words left" are from currentIndex + 1 to end
        const remainingCount = words.length - (currentIndex + 1);
        
        if (remainingCount > 0) {
            // Insert at random position among remaining
            const offset = Math.floor(Math.random() * remainingCount) + 1;
            const targetIndex = currentIndex + offset;
            
            // Insert a copy at the target index
            words.splice(targetIndex, 0, { ...currentWord });
            console.log(`Word "${currentWord.Word}" bookmarked. Re-inserted at index ${targetIndex}.`);
        } else {
            // If it's the last word, insert it somewhere earlier or loop it back
            // For now, let's just push it to the end so it loops back eventually
            words.push({ ...currentWord });
            console.log(`Word "${currentWord.Word}" bookmarked. Added to the end (looping).`);
        }
    } else {
        // If un-bookmarking, we could technically remove future instances
        // but for simplicity and following the specific request "bookmark... and it should be shown again",
        // we'll just keep the logic focused on the "shown again" part.
    }

    // Save state to local storage
    localStorage.setItem('wordwise_data', JSON.stringify(fullWordList));
    
    updateCard();
}

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        currentWorkbook = XLSX.read(data, { type: 'array' });
        
        // Populate sheet selector
        sheetSelect.innerHTML = "";
        currentWorkbook.SheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });
        
        sheetSelectorGroup.style.display = 'block';
        
        // Load first sheet by default
        loadSheetData(currentWorkbook.SheetNames[0]);
    };
    reader.readAsArrayBuffer(file);
}

function loadSheetData(sheetName) {
    if (!currentWorkbook) return;
    
    const worksheet = currentWorkbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet);
    
    if (json.length > 0) {
        fullWordList = json;
        words = shuffle(json);
        // Save to local storage for persistence
        localStorage.setItem('wordwise_data', JSON.stringify(json));
        
        currentIndex = 0;
        updateCard();
        alert(`Loaded ${json.length} words from sheet "${sheetName}" successfully!`);
    } else {
        alert(`No words found in sheet "${sheetName}".`);
    }
}

function handleSheetChange(event) {
    const selectedSheet = event.target.value;
    loadSheetData(selectedSheet);
}

function loadRange() {
    const from = parseInt(rangeFromInput.value);
    const to = parseInt(rangeToInput.value);

    if (isNaN(from) || isNaN(to)) {
        alert("Please enter both From and To numbers.");
        return;
    }

    if (from < 1 || to > fullWordList.length || from > to) {
        alert(`Please enter a valid range between 1 and ${fullWordList.length}.`);
        return;
    }

    // Slice is 0-indexed, users usually think 1-indexed
    const subset = fullWordList.slice(from - 1, to);
    words = shuffle(subset);
    currentIndex = 0;
    updateCard();
    alert(`Loaded ${words.length} words from range ${from}-${to} in jumbled order!`);
}

// Event Listeners
flipBtn.addEventListener('click', toggleFlip);
card.addEventListener('click', toggleFlip);
nextBtn.addEventListener('click', nextWord);
prevBtn.addEventListener('click', prevWord);
shuffleBtn.addEventListener('click', () => {
    words = shuffle(words);
    currentIndex = 0;
    updateCard();
});
excelUpload.addEventListener('change', handleFileUpload);
sheetSelect.addEventListener('change', handleSheetChange);
loadRangeBtn.addEventListener('click', loadRange);
bookmarkBtnFront.addEventListener('click', bookmarkWord);
bookmarkBtnBack.addEventListener('click', bookmarkWord);

resetBtn.addEventListener('click', () => {
    if (confirm("Are you sure you want to clear saved data and reset?")) {
        localStorage.removeItem('wordwise_data');
        fullWordList = [...INITIAL_DATA];
        words = [...INITIAL_DATA];
        currentIndex = 0;
        sheetSelectorGroup.style.display = 'none';
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
            const parsed = JSON.parse(savedData);
            fullWordList = parsed;
            words = shuffle(parsed);
            console.log('Loaded words from localStorage');
        } catch (e) {
            console.error('Failed to parse saved data', e);
            fullWordList = [...INITIAL_DATA];
            words = [...INITIAL_DATA];
        }
    } else {
        fullWordList = [...INITIAL_DATA];
        words = [...INITIAL_DATA];
    }
    updateCard();
}

init();
