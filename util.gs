// util.gs - Utility functions for cosine similarity and K-means clustering

/**
 * Calculate cosine similarity between two vectors
 * @param {Array} vectorA - First vector
 * @param {Array} vectorB - Second vector  
 * @return {number} Cosine similarity value between -1 and 1
 */
function cosineSimilarity(vectorA, vectorB) {
  if (!vectorA || !vectorB || vectorA.length !== vectorB.length) {
    return 0;
  }
  
  var dotProduct = 0;
  var normA = 0;
  var normB = 0;
  
  for (var i = 0; i < vectorA.length; i++) {
    dotProduct += vectorA[i] * vectorB[i];
    normA += vectorA[i] * vectorA[i];
    normB += vectorB[i] * vectorB[i];
  }
  
  if (normA === 0 || normB === 0) {
    return 0;
  }
  
  return dotProduct / (Math.sqrt(normA) * Math.sqrt(normB));
}

/**
 * Perform K-means clustering on embeddings
 * @param {Array} embeddings - Array of embedding vectors
 * @param {number} k - Number of clusters
 * @return {Array} Array of cluster assignments for each embedding
 */
function kMeansClustering(embeddings, k) {
  if (!embeddings || embeddings.length === 0 || k <= 0) {
    return [];
  }
  
  if (k >= embeddings.length) {
    // Each point gets its own cluster
    var assignments = [];
    for (var i = 0; i < embeddings.length; i++) {
      assignments.push(i);
    }
    return assignments;
  }
  
  var dimensions = embeddings[0].length;
  var maxIterations = 20;
  
  // Initialize centroids randomly
  var centroids = [];
  for (var c = 0; c < k; c++) {
    var centroid = [];
    for (var d = 0; d < dimensions; d++) {
      centroid.push(Math.random() * 2 - 1); // Random value between -1 and 1
    }
    centroids.push(centroid);
  }
  
  var assignments = new Array(embeddings.length);
  var converged = false;
  
  for (var iter = 0; iter < maxIterations && !converged; iter++) {
    var newAssignments = new Array(embeddings.length);
    
    // Assign each point to nearest centroid
    for (var p = 0; p < embeddings.length; p++) {
      var bestCluster = 0;
      var bestDistance = euclideanDistance(embeddings[p], centroids[0]);
      
      for (var cl = 1; cl < k; cl++) {
        var distance = euclideanDistance(embeddings[p], centroids[cl]);
        if (distance < bestDistance) {
          bestDistance = distance;
          bestCluster = cl;
        }
      }
      
      newAssignments[p] = bestCluster;
    }
    
    // Check for convergence
    converged = true;
    if (assignments.length > 0) {
      for (var a = 0; a < newAssignments.length; a++) {
        if (newAssignments[a] !== assignments[a]) {
          converged = false;
          break;
        }
      }
    } else {
      converged = false;
    }
    
    assignments = newAssignments;
    
    if (!converged) {
      // Update centroids
      var newCentroids = [];
      for (var cc = 0; cc < k; cc++) {
        var clusterPoints = [];
        for (var pp = 0; pp < embeddings.length; pp++) {
          if (assignments[pp] === cc) {
            clusterPoints.push(embeddings[pp]);
          }
        }
        
        if (clusterPoints.length > 0) {
          newCentroids.push(calculateCentroid(clusterPoints));
        } else {
          // Keep old centroid if no points assigned
          newCentroids.push(centroids[cc]);
        }
      }
      centroids = newCentroids;
    }
  }
  
  return assignments;
}

/**
 * Calculate Euclidean distance between two vectors
 * @param {Array} vectorA - First vector
 * @param {Array} vectorB - Second vector
 * @return {number} Euclidean distance
 */
function euclideanDistance(vectorA, vectorB) {
  if (!vectorA || !vectorB || vectorA.length !== vectorB.length) {
    return Infinity;
  }
  
  var sum = 0;
  for (var i = 0; i < vectorA.length; i++) {
    var diff = vectorA[i] - vectorB[i];
    sum += diff * diff;
  }
  
  return Math.sqrt(sum);
}

/**
 * Calculate centroid (mean) of a set of vectors
 * @param {Array} vectors - Array of vectors
 * @return {Array} Centroid vector
 */
function calculateCentroid(vectors) {
  if (!vectors || vectors.length === 0) {
    return [];
  }
  
  var dimensions = vectors[0].length;
  var centroid = new Array(dimensions);
  
  // Initialize centroid
  for (var d = 0; d < dimensions; d++) {
    centroid[d] = 0;
  }
  
  // Sum all vectors
  for (var v = 0; v < vectors.length; v++) {
    for (var dd = 0; dd < dimensions; dd++) {
      centroid[dd] += vectors[v][dd];
    }
  }
  
  // Calculate mean
  for (var ddd = 0; ddd < dimensions; ddd++) {
    centroid[ddd] /= vectors.length;
  }
  
  return centroid;
}