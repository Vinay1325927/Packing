def test_front_window_displays_joke(client):
    """
    GIVEN a logged-in teller with one submitted joke
    WHEN the teller requests the front window
    THEN the page returns 200 and displays the joke

    LEVEL: HAPPY PATH
    """
    credentials = {
        "username": "testuser",
        "password": "testpass123",
    }
    joke = "Why did the developer go broke? They used up all their cache."

    register_response = client.post(
        "/auth/register",
        json=credentials,
    )
    assert register_response.status_code == 201

    login_response = client.post(
        "/auth/login",
        json=credentials,
    )
    assert login_response.status_code == 200

    create_response = client.post(
        "/jokes",
        json={"text": joke},
    )
    assert create_response.status_code == 201

    response = client.get("/")

    assert response.status_code == 200
    assert joke.encode() in response.data
