import streamlit as st
import pandas as pd
import re
from groq import Groq
import io
from datetime import datetime
import time

# Page configuration
st.set_page_config(
    page_title="Orange Maroc - Review Topic Classifier",
    page_icon="🍊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for Orange branding
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #FF6600 0%, #FF8533 100%);
        padding: 1rem;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .main-header h1 {
        color: white;
        margin: 0;
        text-align: center;
    }
    .main-header p {
        color: white;
        margin: 0.5rem 0 0 0;
        text-align: center;
        opacity: 0.9;
    }
    .stButton > button {
        background-color: #FF6600;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        font-weight: bold;
    }
    .stButton > button:hover {
        background-color: #E55A00;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

class TopicGenerator:
    def __init__(self, model_name="llama-3.3-70b-versatile"):
        self.client = Groq(api_key="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")
        self.model_name = model_name
    
    def generate(self, review_text):
        # Remove translation markers
        clean_review = review_text.replace("(Translated by Google)", "").replace("(Original)", "").strip()
        
        prompt = f"""
Vous êtes un expert en catégorisation des avis clients pour les services télécoms d'Orange Maroc.

Analysez l'avis client suivant et classez-le dans UNE SEULE des catégories ci-dessous, en vous basant sur le SUJET PRINCIPAL abordé :

- service client : avis sur la qualité de l'assistance, la résolution des problèmes, la réactivité du service
- comportement du personnel : attitude des employés, politesse, professionnalisme, amabilité
- qualité du réseau : force du signal, qualité des appels, couverture, vitesse internet
- tarification : coûts, tarifs, perception de cherté ou bon rapport qualité/prix
- installation : mise en place des services, visite de technicien, activation
- assistance technique : dépannage, réparation, résolution de problèmes techniques
- expérience en boutique : conditions d’accueil, propreté, accessibilité, file d’attente
- qualité des produits : fiabilité des téléphones, équipements, cartes SIM, appareils
- facturation : erreurs de facturation, paiements, méthodes de règlement, gestion de compte
- temps d’attente : attente en boutique, prise de rendez-vous, files d'attente
- satisfaction générale : avis global sans focus précis, remerciements, recommandations

Instructions :
- Répondez UNIQUEMENT par le nom de la catégorie
- Utilisez des minuscules
- Si le sujet est flou, répondez "autre"

Avis : "{clean_review}"
"""

        
        try:
            response = self.client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {"role": "system", "content": "You are an expert in categorizing customer reviews."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=20,
                temperature=0.2,
            )
            
            # Extract and clean the topic
            topic = response.choices[0].message.content.strip()
            
            # Validate and standardize topic
        
            topics_mapping = {
    "service client": "Service Client",
    "comportement du personnel": "Comportement du Personnel",
    "qualité du réseau": "Qualité du Réseau",
    "tarification": "Tarification",
    "installation": "Installation",
    "assistance technique": "Assistance Technique",
    "expérience en boutique": "Expérience en Boutique",
    "qualité des produits": "Qualité des Produits",
    "facturation": "Facturation",
    "temps d’attente": "Temps d’Attente",
    "satisfaction générale": "Satisfaction Générale"
}

            
            # Find the closest match or use the original if no match
            for key, value in topics_mapping.items():
                if key in topic.lower():
                    return value
            
            return "Other"
            
        except Exception as e:
            st.error(f"Error in topic extraction: {e}")
            return "Other"

def clean_reviews(uploaded_file):
    """
    Load and clean the Excel file containing customer reviews
    
    Args:
        uploaded_file: Streamlit uploaded file object
        
    Returns:
        pandas.DataFrame: Cleaned DataFrame
    """
    try:
        # Load Excel file
        df = pd.read_excel(uploaded_file, header=1)  # Skip first row
        
        # Keep first 18 columns
        if df.shape[1] >= 18:
            df = df.iloc[:, :18]
        
        # Ensure column names are properly formatted
        df.columns = df.columns.str.strip()
        
        # Rename columns if needed (for consistency)
        column_mapping = {
            'Localité': 'Ville',
            'Store': 'Code magasin',
            'Review': 'content'
        }
        
        # Apply mappings for columns that exist
        for old_col, new_col in column_mapping.items():
            if old_col in df.columns:
                df.rename(columns={old_col: new_col}, inplace=True)
        
        # Ensure review content column exists
        if 'content' not in df.columns:
            # Try to find a column containing review text
            for col in df.columns:
                if col.lower().find('review') >= 0 or col.lower().find('avis') >= 0:
                    df.rename(columns={col: 'content'}, inplace=True)
                    break
        
        # Clean the review content
        if 'content' in df.columns:
            # Make sure content is string
            df["content"] = df["content"].astype(str).str.strip()
            
            # Remove empty reviews
            df = df[df["content"].str.len() > 0]
            df = df[df["content"] != "nan"]
            
            # Clean translation markers
            def clean_review_text(text):
                # Remove translation markers
                text = re.sub(r'\(Translated by Google\)', '', text)
                text = re.sub(r'\(Original\)', ' | Original: ', text)
                return text.strip()
            
            df["content"] = df["content"].apply(clean_review_text)
            
            # Create a 'language' column (optional)
            def detect_language(text):
                text = text.lower()
                if '| original:' in text:
                    return 'Translated'
                elif any(word in text for word in ['merci', 'bonjour', 'service']):
                    return 'French'
                else:
                    return 'Other'
            
            df['language'] = df['content'].apply(detect_language)
        
        return df
    
    except Exception as e:
        raise Exception(f"Error cleaning reviews: {e}")

def process_reviews_with_topics(df, progress_bar, status_text):
    """Process reviews and add topic classification"""
    generator = TopicGenerator()
    topics = []
    
    total_reviews = len(df)
    
    for i, review in enumerate(df['content']):
        # Update progress
        progress = (i + 1) / total_reviews
        progress_bar.progress(progress)
        status_text.text(f"Processing review {i + 1} of {total_reviews}...")
        
        # Generate topic
        topic = generator.generate(review)
        topics.append(topic)
        
        # Small delay to prevent API rate limiting
        time.sleep(0.1)
    
    # Add topics to dataframe
    df['Topic'] = topics
    return df

def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>🍊 Orange Maroc - Review Topic Classifier</h1>
        <p>Customer Experience & Digitalisation Team | Automated Topic Analysis Tool</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Introduction
    st.markdown("""
    <div class="info-box">
        <h3>📋 How to use this tool:</h3>
        <ol>
            <li><strong>Upload</strong> your Excel file containing Google My Business reviews</li>
            <li><strong>Preview</strong> the cleaned data to ensure it looks correct</li>
            <li><strong>Process</strong> the reviews to automatically classify topics using AI</li>
            <li><strong>Download</strong> the enriched Excel file for use in Power BI</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)
    
    # File upload section
    st.header("📁 Step 1: Upload Your Excel File")
    uploaded_file = st.file_uploader(
        "Choose your Excel file containing Google My Business reviews",
        type=['xlsx', 'xls'],
        help="Make sure your file contains columns for store name, city, and review content"
    )
    
    if uploaded_file is not None:
        try:
            # Clean the data
            with st.spinner("🧹 Cleaning and preparing your data..."):
                df_cleaned = clean_reviews(uploaded_file)
            
            st.success(f"✅ File successfully loaded and cleaned! Found {len(df_cleaned)} reviews.")
            
            # Data preview section
            st.header("👀 Step 2: Preview Your Data")
            
            # Show basic statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Reviews", len(df_cleaned))
            with col2:
                if 'Ville' in df_cleaned.columns:
                    st.metric("Cities", df_cleaned['Ville'].nunique())
                else:
                    st.metric("Cities", "N/A")
            with col3:
                if 'Code magasin' in df_cleaned.columns:
                    st.metric("Stores", df_cleaned['Code magasin'].nunique())
                else:
                    st.metric("Stores", "N/A")
            
            # Show column information
            st.subheader("📊 Data Structure")
            st.write(f"**Columns found:** {', '.join(df_cleaned.columns.tolist())}")
            
            # Show preview of data
            st.subheader("🔍 Data Preview (First 5 rows)")
            st.dataframe(df_cleaned.head(), use_container_width=True)
            
            # Processing section
            st.header("🤖 Step 3: AI Topic Classification")
            
            if 'content' not in df_cleaned.columns:
                st.error("❌ No review content column found. Please make sure your Excel file contains a column with review text.")
                return
            
            
            
            if st.button("🚀 Start Topic Classification", type="primary"):
                # Create progress indicators
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    # Process reviews
                    df_with_topics = process_reviews_with_topics(df_cleaned, progress_bar, status_text)
                    
                    # Clear progress indicators
                    progress_bar.empty()
                    status_text.empty()
                    
                    st.success("🎉 Topic classification completed successfully!")
                    
                    # Show results
                    st.header("📈 Step 4: Results & Download")
                    
                    # Topic distribution
                    st.subheader("📊 Topic Distribution")
                    topic_counts = df_with_topics['Topic'].value_counts()
                    
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        st.bar_chart(topic_counts)
                    with col2:
                        st.dataframe(topic_counts.reset_index().rename(columns={'index': 'Topic', 'Topic': 'Count'}))
                    
                    # Show processed data preview
                    st.subheader("🔍 Processed Data Preview")
                    st.dataframe(df_with_topics.head(10), use_container_width=True)
                    
                    # Download section
                    st.subheader("💾 Download Your Results")
                    
                    # Prepare Excel file for download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_with_topics.to_excel(writer, sheet_name='Reviews_with_Topics', index=False)
                    
                    output.seek(0)
                    
                    # Generate filename with timestamp
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"Orange_Reviews_Topics_{timestamp}.xlsx"
                    
                    st.download_button(
                        label="📥 Download Enriched Excel File",
                        data=output.getvalue(),
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                    st.markdown("""
                    <div class="success-box">
                        <h4>✅ Next Steps:</h4>
                        <ol>
                            <li>Download the enriched Excel file using the button above</li>
                            <li>Save it to your Power BI data source location</li>
                            <li>Open Power BI and click "Refresh" to update your dashboard</li>
                            <li>Your visualizations will now show the new topic classifications!</li>
                        </ol>
                    </div>
                    """, unsafe_allow_html=True)
                    
                except Exception as e:
                    st.error(f"❌ Error during topic classification: {str(e)}")
                    st.info("💡 Please check your internet connection and try again. If the problem persists, contact the development team.")
        
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.info("💡 Please make sure your Excel file is properly formatted and try again.")
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>🍊 Orange Maroc - Customer Experience & Digitalisation Team</p>
        <p><small>Powered by Groq AI & LLaMA 3.3 70B | Built with Streamlit</small></p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()