def main():
    luna = LunaAssistant()
    luna.wish_me()
    
    running = True
    while running:
        query = luna.take_command()
        running = luna.process_command(query)

if __name__ == "__main__":
    main()