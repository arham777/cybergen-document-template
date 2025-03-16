import { NextResponse } from 'next/server';

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    
    // Forward the request to our Python backend
    const response = await fetch(`${process.env.NEXT_PUBLIC_API_URL || ''}/api/generate-from-file`, {
      method: 'POST',
      body: formData,
    });
    
    const data = await response.json();
    
    return NextResponse.json(data);
  } catch (error) {
    console.error('Error generating document:', error);
    return NextResponse.json(
      { success: false, error: 'Internal server error' },
      { status: 500 }
    );
  }
} 