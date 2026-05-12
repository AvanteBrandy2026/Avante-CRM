import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://xsuvahsnaulfopdgxqpp.supabase.co'
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhzdXZhaHNuYXVsZm9wZGd4cXBwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg0OTg0MDIsImV4cCI6MjA5NDA3NDQwMn0.XbTqvGJaOa8UQWuzJfXj0vlF82u9U_q0-YUIybiCQbI'

export const supabase = createClient(supabaseUrl, supabaseKey)
