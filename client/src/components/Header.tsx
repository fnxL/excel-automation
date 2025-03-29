import * as React from 'react'

function Header({ children }: { children: React.ReactNode }) {
    return (
        <header className="container mx-auto px-4 py-8">
            <h1 className="text-4xl font-extrabold tracking-tight lg:text-5xl text-center">
                {children}
            </h1>
        </header>
    )
}

export default Header
