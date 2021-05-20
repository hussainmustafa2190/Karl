import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.decomposition import TruncatedSVD
from sklearn.neighbors import BallTree
from sklearn.base import BaseEstimator
from sklearn.pipeline import make_pipeline
import warnings
warnings.filterwarnings("ignore")
import pickle 
import joblib


lines = pd.read_csv("C:/topical_chat.csv")
lines = lines.iloc[:,1].apply(lambda x: x.lower())
subtitles = pd.DataFrame(columns=['context', 'reply'])
subtitles['context'] = lines
subtitles['reply'] = lines[1:].reset_index(drop = True) + ['...']




"""subtitles = pd.DataFrame(columns=['context', 'reply'])
subtitles['context'] = lines
subtitles['context'] = subtitles['context'].apply(lambda x: x.lower())
subtitles['reply'] = lines[1:] + ['...']
subtitles['reply'] = subtitles['reply'].apply(lambda x: x.lower())"""




"""for sign in ['!', '?', ',', '.', ':']:
    subtitles['context'] = subtitles['context'].apply(lambda x: x.replace(sign, f' {sign}'))
    subtitles['reply'] = subtitles['reply'].apply(lambda x: x.replace(sign, f' {sign}'))"""
    
subtitles.info()
subtitles.iloc[100:120]


vectorizer = TfidfVectorizer()
vectorizer.fit(subtitles.context)

matrix_big = vectorizer.transform(subtitles.context)
matrix_big.shape


svd = TruncatedSVD(n_components=300, algorithm='arpack')

svd.fit(matrix_big)
matrix_small = svd.transform(matrix_big)





def softmax(x):
    proba = np.exp(-x)
    return proba/sum(proba)


class NeighborSampler(BaseEstimator):
    def __init__(self, k=5, temperature = 1.0):
        self.k = k
        self.temperature = temperature
    
    def fit(self, X, y):
        self.tree_ = BallTree(X)
        self.y_ = np.array(y)
        
    def predict(self, X, random_state = None):
        distances, indeces = self.tree_.query(X, return_distance = True, k = self.k)
        result = []
        for distance, index in zip(distances, indeces):
            result.append(np.random.choice(index, p = softmax(distance * self.temperature)))
            
        return self.y_[result]
    
    
    
ns = NeighborSampler()
ns.fit(matrix_small, subtitles.reply)


pipe = make_pipeline(vectorizer, svd, ns)


            
            
if __name__=="__main__":
    pipe.__module__ = "temp1"
        
    with open("pipe","wb") as f :
        pickle.dump(pipe,f)



    
    
    """
    
    
joblib.dump(pipe,"pipe_model.pkl")

"""
