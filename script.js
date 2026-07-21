// Interactive Scripts for Engr. Muhammad Waqas Portfolio

document.addEventListener('DOMContentLoaded', () => {
    // 1. Navbar Scroll Effect
    const navbar = document.getElementById('navbar');
    window.addEventListener('scroll', () => {
        if (window.scrollY > 50) {
            navbar.style.background = 'rgba(11, 15, 25, 0.95)';
            navbar.style.boxShadow = '0 4px 20px rgba(0, 0, 0, 0.5)';
        } else {
            navbar.style.background = 'rgba(11, 15, 25, 0.85)';
            navbar.style.boxShadow = 'none';
        }
    });

    // 2. Mobile Nav Toggle
    const navToggle = document.getElementById('nav-toggle');
    const navMenu = document.getElementById('nav-menu');
    if (navToggle && navMenu) {
        navToggle.addEventListener('click', () => {
            navMenu.classList.toggle('active');
        });
    }
});

// 3. Project Tab Filter Logic
function filterProjects(category) {
    const buttons = document.querySelectorAll('.tab-btn');
    buttons.forEach(btn => btn.classList.remove('active'));

    // Set active button
    event.currentTarget.classList.add('active');

    const cards = document.querySelectorAll('.project-card');
    cards.forEach(card => {
        const cat = card.getAttribute('data-category');
        if (category === 'all' || cat === category) {
            card.style.display = 'flex';
            card.style.opacity = '1';
        } else {
            card.style.display = 'none';
            card.style.opacity = '0';
        }
    });
}

// 4. Modal Data & Logic
const modalData = {
    'ai-architect': {
        title: 'AI-Architect — Generative AI Layout & BIM Engine',
        subtitle: 'Generative Architecture & CAD/BIM Automation Platform',
        github: 'https://github.com/engrwaqas034-AI/AI-Architect',
        details: `
            <p><strong>Core Concept:</strong> AI-Architect is an advanced generative design system built with Python, spatial Constraint Satisfaction Problem (CSP) algorithms, and parametric SVG/BIM export modules.</p>
            <br>
            <h4>Key Features & Architecture:</h4>
            <ul>
                <li><strong>Plot & Setback Calculation:</strong> Computes mandatory setback boundaries according to local building authority bylaws.</li>
                <li><strong>Generative Layout Engine:</strong> Places living rooms, bedrooms, kitchens, and circulation corridors with zero room overlapping and optimal natural light orientation.</li>
                <li><strong>Multi-Floor & CAD Exporter:</strong> Generates multi-floor plans and exports vector-accurate DXF / IFC coordinate structures for direct import into AutoCAD & Revit.</li>
            </ul>
        `
    },
    'claim-analyzer': {
        title: 'Construction_Claim_Analyzer — LLM & RAG Claim System',
        subtitle: 'AI-Powered FIDIC Contract & Delay Analysis',
        github: 'https://github.com/engrwaqas034-AI/Construction_Claim_Analyzer',
        details: `
            <p><strong>Core Concept:</strong> A specialized LLM & Retrieval-Augmented Generation (RAG) platform that audits construction project delay claims, Extension of Time (EOT) applications, and FIDIC contract disputes.</p>
            <br>
            <h4>Key Features & Architecture:</h4>
            <ul>
                <li><strong>Contractual Vector DB:</strong> Embeds FIDIC General Conditions of Contract, project baseline schedules (Primavera P6), daily site logs, and IPCs into a semantic search index.</li>
                <li><strong>Delay Attribution Matrix:</strong> Automatically identifies critical path delays, contractor/client risk allocation, and generates legal/engineering risk scores.</li>
                <li><strong>Rapid Audit:</strong> Audits 500+ page claim submissions in minutes, outputting structured executive summaries and vulnerability reports.</li>
            </ul>
        `
    },
    'boq-generator': {
        title: 'AI BOQ & Quantity Surveying Generator',
        subtitle: 'Automated BOQ, BBS & Cost Estimation Pipeline',
        github: 'https://github.com/engrwaqas034-AI',
        details: `
            <p><strong>Core Concept:</strong> An automated AI quantity surveying pipeline that calculates concrete volumes, steel reinforcement Bar Bending Schedules (BBS), and material take-offs directly from BIM schemas and structural parameters.</p>
            <br>
            <h4>Key Features & Architecture:</h4>
            <ul>
                <li><strong>Automated Take-Off:</strong> Parses structural drawings to compute foundation, column, beam, and slab quantities automatically.</li>
                <li><strong>Revit API Integration:</strong> Connects directly with Autodesk Revit models to extract live material schedules.</li>
                <li><strong>Excel & Cost Database Linkage:</strong> Outputs formatted Bills of Quantities (BOQs) linked with unit rate cost databases for rapid tender preparation.</li>
            </ul>
        `
    }
};

function openProjectModal(projectId) {
    const modal = document.getElementById('project-modal');
    const modalContent = document.getElementById('modal-content');
    const data = modalData[projectId];

    if (data && modal && modalContent) {
        modalContent.innerHTML = `
            <div style="color: var(--primary-light); font-size: 0.85rem; font-weight: 700; text-transform: uppercase; margin-bottom: 0.5rem;">${data.subtitle}</div>
            <h2 style="font-family: var(--font-heading); font-size: 1.6rem; color: var(--text-main); margin-bottom: 1rem;">${data.title}</h2>
            <div style="font-size: 0.95rem; color: var(--text-muted); line-height: 1.6;">${data.details}</div>
            <div style="margin-top: 2rem;">
                <a href="${data.github}" target="_blank" class="btn btn-primary">
                    <i class="fa-brands fa-github"></i> Open Repository on GitHub
                </a>
            </div>
        `;
        modal.classList.add('active');
    }
}

function closeProjectModal() {
    const modal = document.getElementById('project-modal');
    if (modal) {
        modal.classList.remove('active');
    }
}

// Close modal when clicking outside
window.addEventListener('click', (e) => {
    const modal = document.getElementById('project-modal');
    if (e.target === modal) {
        closeProjectModal();
    }
});
